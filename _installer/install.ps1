param(
    [Parameter(Mandatory=$true)][string]$PackageRoot,
    [string]$InstallMode = "",
    [string]$SharedDataPath = "",
    [string]$ArchiveRootPath = "",
    [string]$TrabajosYearPath = "",
    [string]$TargetMailbox = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Get-EnvValue([string]$Path, [string]$Name) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return "" }
    $Pattern = '^\s*' + [Regex]::Escape($Name) + '\s*=\s*(.*)$'
    foreach ($Line in Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue) {
        if ($Line -match $Pattern) { return $Matches[1].Trim().Trim('"').Trim("'") }
    }
    return ""
}

function Set-EnvValue([string]$Path, [string]$Name, [string]$Value) {
    $Lines = @()
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        $Lines = @(Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue)
    }
    $Pattern = '^\s*' + [Regex]::Escape($Name) + '\s*='
    $Found = $false
    for ($Index = 0; $Index -lt $Lines.Count; $Index++) {
        if ($Lines[$Index] -match $Pattern) {
            $Lines[$Index] = "$Name=$Value"
            $Found = $true
        }
    }
    if (-not $Found) {
        if ($Lines.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($Lines[-1])) { $Lines += "" }
        $Lines += "$Name=$Value"
    }
    Set-Content -LiteralPath $Path -Value $Lines -Encoding UTF8
}

function Normalize-InstallerPath([string]$Path) {
    $Path = $Path.Trim().Trim('"').Trim("'")
    try { return [System.IO.Path]::GetFullPath($Path) }
    catch { throw "The installer folder path is invalid: $Path" }
}

function Test-ProcessorWriteAccess([string]$Path) {
    $Probe = Join-Path $Path ((".ingevia_write_test_{0}.tmp") -f [Guid]::NewGuid().ToString("N"))
    try {
        "test" | Set-Content -LiteralPath $Probe -Encoding ASCII
        Remove-Item -LiteralPath $Probe -Force
    } catch {
        throw "The processing computer cannot write to the shared data folder: $Path. Check network access and permissions. Details: $($_.Exception.Message)"
    }
}

function Get-ValidCachedFile([string]$Path, [long]$MinimumBytes) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try { return ((Get-Item -LiteralPath $Path).Length -ge $MinimumBytes) }
    catch { return $false }
}

function Download-RequiredFile([string]$Uri, [string]$Destination, [long]$MinimumBytes, [string]$Description) {
    if (Get-ValidCachedFile $Destination $MinimumBytes) { return }
    Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
    Write-Host "      Downloading $Description..."
    try {
        Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing
    } catch {
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        throw "$Description could not be downloaded. Connect this PC to the internet, or run PREPARE_OFFLINE_COPY.bat on another Windows PC and copy the completed installer folder. Details: $($_.Exception.Message)"
    }
    if (-not (Get-ValidCachedFile $Destination $MinimumBytes)) {
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        throw "The downloaded $Description file is incomplete or invalid. Try again on another internet connection, or prepare an offline copy."
    }
}

function Enable-EmbeddedPythonSite([string]$Root) {
    $PthFile = Get-ChildItem -LiteralPath $Root -Filter "python*._pth" -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $PthFile) { throw "The portable Python configuration file was not found after extraction." }

    $Lines = @(Get-Content -LiteralPath $PthFile.FullName)
    $Updated = @()
    $HasSitePackages = $false
    $HasImportSite = $false

    foreach ($Line in $Lines) {
        $Trimmed = $Line.Trim()
        if ($Trimmed -match '^#\s*import\s+site\s*$' -or $Trimmed -match '^import\s+site\s*$') {
            if (-not $HasImportSite) {
                $Updated += 'import site'
                $HasImportSite = $true
            }
            continue
        }
        if ($Trimmed -ieq 'Lib\site-packages' -or $Trimmed -ieq 'Lib/site-packages') {
            if (-not $HasSitePackages) {
                $Updated += 'Lib\site-packages'
                $HasSitePackages = $true
            }
            continue
        }
        $Updated += $Line
    }
    if (-not $HasSitePackages) { $Updated += 'Lib\site-packages' }
    if (-not $HasImportSite) { $Updated += 'import site' }

    Set-Content -LiteralPath $PthFile.FullName -Value $Updated -Encoding ASCII
    New-Item -ItemType Directory -Force -Path (Join-Path $Root 'Lib\site-packages') | Out-Null
}

function Install-PortablePythonRuntime(
    [string]$RuntimeRoot,
    [string]$RuntimeArchive,
    [string]$RuntimeArchiveUrl,
    [string]$GetPipScript,
    [string]$GetPipUrl
) {
    $PythonExe = Join-Path $RuntimeRoot 'python.exe'
    $PythonWExe = Join-Path $RuntimeRoot 'pythonw.exe'

    if ((Test-Path -LiteralPath $PythonExe -PathType Leaf) -and (Test-Path -LiteralPath $PythonWExe -PathType Leaf)) {
        return
    }

    Download-RequiredFile $RuntimeArchiveUrl $RuntimeArchive 8000000 'official portable Python runtime'
    Download-RequiredFile $GetPipUrl $GetPipScript 500000 'official pip bootstrap script'

    Remove-Item -LiteralPath $RuntimeRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $RuntimeRoot | Out-Null

    try {
        Expand-Archive -LiteralPath $RuntimeArchive -DestinationPath $RuntimeRoot -Force
    } catch {
        Remove-Item -LiteralPath $RuntimeRoot -Recurse -Force -ErrorAction SilentlyContinue
        throw "The portable Python runtime could not be extracted. The cached archive may be damaged. Delete the _offline_cache folder and run the installer again. Details: $($_.Exception.Message)"
    }

    Enable-EmbeddedPythonSite $RuntimeRoot

    if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf) -or -not (Test-Path -LiteralPath $PythonWExe -PathType Leaf)) {
        throw "The portable Python runtime was extracted, but python.exe or pythonw.exe is missing."
    }

    $BootstrapArguments = @($GetPipScript, '--disable-pip-version-check', 'pip<27')
    & $PythonExe @BootstrapArguments
    if ($LASTEXITCODE -ne 0) {
        throw "pip could not be installed into the private portable Python runtime."
    }
}

$PackageRoot = Normalize-InstallerPath $PackageRoot
if (-not (Test-Path -LiteralPath $PackageRoot -PathType Container)) {
    throw "The installer folder could not be found: $PackageRoot"
}
$PackageRoot = (Get-Item -LiteralPath $PackageRoot).FullName
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$InstallMode = $InstallMode.Trim().ToUpperInvariant()
if ($InstallMode -notin @("PROCESSOR", "VIEWER")) {
    do {
        $Choice = (Read-Host "Enter 1 for PROCESSING COMPUTER or 2 for VIEWER COMPUTER").Trim()
    } until ($Choice -in @("1", "2"))
    $InstallMode = if ($Choice -eq "1") { "PROCESSOR" } else { "VIEWER" }
}

$ProductName = "INGEVIA Email Assistant"
$InstallRoot = Join-Path $env:LOCALAPPDATA $ProductName
$PayloadRoot = Join-Path $PackageRoot "payload"
$AppRoot = Join-Path $InstallRoot "app"
$RuntimeRoot = Join-Path $InstallRoot "runtime"
$MaintenanceRoot = Join-Path $InstallRoot "maintenance"
$CacheRoot = Join-Path $PackageRoot "_offline_cache"
$PythonArchiveName = "python-3.11.9-embed-amd64.zip"
$PythonArchiveUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
$PythonArchive = Join-Path $CacheRoot $PythonArchiveName
$GetPipUrl = "https://bootstrap.pypa.io/get-pip.py"
$GetPipScript = Join-Path $CacheRoot "get-pip.py"
$Requirements = Join-Path $PayloadRoot "requirements-runtime.txt"
$WheelsRoot = Join-Path $CacheRoot "wheels"
$OfflineMarker = Join-Path $CacheRoot "CACHE_COMPLETE.txt"
$ExistingEnv = Join-Path $AppRoot ".env"

if (-not [Environment]::Is64BitOperatingSystem) { throw "This package supports 64-bit Windows only." }
if (-not (Test-Path -LiteralPath $PayloadRoot -PathType Container)) { throw "Installer payload is missing: $PayloadRoot" }

$PreviousSharedPath = Get-EnvValue $ExistingEnv "SHARED_DATA_PATH"
if ([string]::IsNullOrWhiteSpace($PreviousSharedPath)) { $PreviousSharedPath = Get-EnvValue $ExistingEnv "OUTPUT_ROOT" }
$PreviousArchivePath = Get-EnvValue $ExistingEnv "ARCHIVE_ROOT"
$PreviousTrabajosPath = Get-EnvValue $ExistingEnv "PROJECT_YEAR_ROOT"
$PreviousMailbox = Get-EnvValue $ExistingEnv "TARGET_MAILBOX"
if ([string]::IsNullOrWhiteSpace($PreviousMailbox)) { $PreviousMailbox = Get-EnvValue $ExistingEnv "BOSS_EMAIL" }

Write-Host ""
Write-Host "Installing $ProductName" -ForegroundColor Cyan
Write-Host "Role: $InstallMode"
Write-Host "Destination: $InstallRoot"
Write-Host ""

if ([string]::IsNullOrWhiteSpace($SharedDataPath)) {
    if ($InstallMode -eq "PROCESSOR") {
        $DefaultShared = if ([string]::IsNullOrWhiteSpace($PreviousSharedPath)) { "C:\EmailAssistant\Output" } else { $PreviousSharedPath }
        Write-Host "The processing computer and all viewer PCs must use the SAME data folder." -ForegroundColor Yellow
        Write-Host "For multiple PCs, use a network/UNC path such as: \\OFFICE-SERVER\EmailAssistantData"
        $Entered = Read-Host "Shared data folder [$DefaultShared]"
        $SharedDataPath = if ([string]::IsNullOrWhiteSpace($Entered)) { $DefaultShared } else { $Entered.Trim().Trim('"') }
    } else {
        $DefaultViewer = $PreviousSharedPath
        Write-Host "Enter the shared folder used by the processing computer." -ForegroundColor Yellow
        Write-Host "Example: \\OFFICE-SERVER\EmailAssistantData"
        do {
            $Prompt = if ([string]::IsNullOrWhiteSpace($DefaultViewer)) { "Shared data folder" } else { "Shared data folder [$DefaultViewer]" }
            $Entered = Read-Host $Prompt
            $SharedDataPath = if ([string]::IsNullOrWhiteSpace($Entered)) { $DefaultViewer } else { $Entered.Trim().Trim('"') }
        } until (-not [string]::IsNullOrWhiteSpace($SharedDataPath))
    }
}
$SharedDataPath = $SharedDataPath.Trim().Trim('"')

# The user selects the EXACT existing TRABAJOS folder used by the engineering team,
# for example P:\TRABAJOS 2026.  Internally the existing routing code still uses
# ARCHIVE_ROOT as its parent (P:\), then appends TRABAJOS <email year>.  Keeping
# PROJECT_YEAR_ROOT separately makes the intended destination explicit and avoids
# accidentally writing into Email Assistant Data\TRABAJOS 2026.
$CurrentYear = (Get-Date).Year
$ExpectedYearFolderName = "TRABAJOS $CurrentYear"

if ([string]::IsNullOrWhiteSpace($TrabajosYearPath)) {
    $DefaultTrabajos = $PreviousTrabajosPath
    if ([string]::IsNullOrWhiteSpace($DefaultTrabajos)) {
        if (-not [string]::IsNullOrWhiteSpace($PreviousArchivePath) -and $PreviousArchivePath -ne $PreviousSharedPath) {
            $ArchiveLeaf = Split-Path -Path $PreviousArchivePath -Leaf
            if ($ArchiveLeaf -match '^TRABAJOS \d{4}$') {
                $DefaultTrabajos = $PreviousArchivePath
            } else {
                $DefaultTrabajos = Join-Path $PreviousArchivePath $ExpectedYearFolderName
            }
        } else {
            $DefaultTrabajos = "P:\TRABAJOS $CurrentYear"
        }
    }

    Write-Host ""
    Write-Host "Enter the EXACT engineering TRABAJOS folder used for processed emails and attachments." -ForegroundColor Yellow
    Write-Host "For this year it should normally be: P:\TRABAJOS $CurrentYear"
    Write-Host "You may also use its UNC equivalent, for example: \\SERVER\PROJECTS\TRABAJOS $CurrentYear"
    Write-Host "Do NOT select Email Assistant Data\TRABAJOS $CurrentYear." -ForegroundColor Yellow
    $EnteredTrabajos = Read-Host "TRABAJOS $CurrentYear folder [$DefaultTrabajos]"
    $TrabajosYearPath = if ([string]::IsNullOrWhiteSpace($EnteredTrabajos)) { $DefaultTrabajos } else { $EnteredTrabajos.Trim().Trim('"') }
}
$TrabajosYearPath = $TrabajosYearPath.Trim().Trim('"')

if (-not (Test-Path -LiteralPath $TrabajosYearPath -PathType Container)) {
    throw "The TRABAJOS folder is unavailable: $TrabajosYearPath. Connect to the office network / P drive and verify the path and permissions."
}

$SelectedLeaf = Split-Path -Path $TrabajosYearPath -Leaf
if ($SelectedLeaf -notmatch '^TRABAJOS \d{4}$') {
    throw "Please select the TRABAJOS year folder itself (for example P:\TRABAJOS $CurrentYear), not its parent and not Email Assistant Data. Selected: $TrabajosYearPath"
}

$ArchiveRootPath = Split-Path -Path $TrabajosYearPath -Parent
if ([string]::IsNullOrWhiteSpace($ArchiveRootPath)) {
    throw "Could not determine the parent archive root from: $TrabajosYearPath"
}

if ($InstallMode -eq "PROCESSOR") {
    if ([string]::IsNullOrWhiteSpace($TargetMailbox)) {
        $DefaultMailbox = if ([string]::IsNullOrWhiteSpace($PreviousMailbox)) { "m.vera@ingevia.com" } else { $PreviousMailbox }
        $EnteredMailbox = Read-Host "Boss Outlook mailbox to process [$DefaultMailbox]"
        $TargetMailbox = if ([string]::IsNullOrWhiteSpace($EnteredMailbox)) { $DefaultMailbox } else { $EnteredMailbox.Trim() }
    }
    try { New-Item -ItemType Directory -Force -Path $SharedDataPath | Out-Null }
    catch { throw "The shared data folder could not be created or opened: $SharedDataPath. Details: $($_.Exception.Message)" }
    Test-ProcessorWriteAccess $SharedDataPath
    # The folder must already be the real engineering archive (e.g. P:\TRABAJOS 2026).
    # We deliberately do not create a second TRABAJOS folder under Email Assistant Data.
    Test-ProcessorWriteAccess $TrabajosYearPath
} else {
    if (-not (Test-Path -LiteralPath $SharedDataPath -PathType Container)) {
        throw "The viewer cannot access the shared data folder: $SharedDataPath. Connect to the office network and verify folder permissions, then run the installer again."
    }
    if ([string]::IsNullOrWhiteSpace($TargetMailbox)) {
        $TargetMailbox = if ([string]::IsNullOrWhiteSpace($PreviousMailbox)) { "m.vera@ingevia.com" } else { $PreviousMailbox }
    }
}

New-Item -ItemType Directory -Force -Path $InstallRoot, $AppRoot, $RuntimeRoot, $MaintenanceRoot, $CacheRoot | Out-Null

# Preserve current settings during an update, then overwrite only role-specific values.
$EnvBackup = Join-Path $env:TEMP (("ingevia_email_assistant_env_backup_{0}.txt") -f [Guid]::NewGuid().ToString("N"))
if (Test-Path -LiteralPath $ExistingEnv) { Copy-Item -LiteralPath $ExistingEnv -Destination $EnvBackup -Force }

Write-Host "[1/6] Copying application files..."
Copy-Item (Join-Path $PayloadRoot "app\*") $AppRoot -Recurse -Force
Copy-Item (Join-Path $PayloadRoot "maintenance\*") $MaintenanceRoot -Recurse -Force
Copy-Item $Requirements (Join-Path $InstallRoot "requirements-runtime.txt") -Force
if (Test-Path -LiteralPath $EnvBackup) {
    Copy-Item -LiteralPath $EnvBackup -Destination $ExistingEnv -Force
    Remove-Item -LiteralPath $EnvBackup -Force
}
if (-not (Test-Path -LiteralPath $ExistingEnv -PathType Leaf)) {
    Copy-Item (Join-Path $AppRoot ".env.example") $ExistingEnv -Force
}

Set-EnvValue $ExistingEnv "APP_MODE" $InstallMode
Set-EnvValue $ExistingEnv "SHARED_DATA_PATH" $SharedDataPath
Set-EnvValue $ExistingEnv "OUTPUT_ROOT" $SharedDataPath
Set-EnvValue $ExistingEnv "ARCHIVE_ROOT" $ArchiveRootPath
Set-EnvValue $ExistingEnv "PROJECT_YEAR_ROOT" $TrabajosYearPath
Set-EnvValue $ExistingEnv "TARGET_MAILBOX" $TargetMailbox
Set-EnvValue $ExistingEnv "BOSS_EMAIL" $TargetMailbox

$PythonExe = Join-Path $RuntimeRoot "python.exe"
$PythonWExe = Join-Path $RuntimeRoot "pythonw.exe"
if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf) -or -not (Test-Path -LiteralPath $PythonWExe -PathType Leaf)) {
    Write-Host "[2/6] Installing the private portable Python runtime..."
    Install-PortablePythonRuntime $RuntimeRoot $PythonArchive $PythonArchiveUrl $GetPipScript $GetPipUrl
} else {
    Write-Host "[2/6] Private Python runtime already present; keeping it."
}

Write-Host "[3/6] Installing application libraries..."
if ((Test-Path -LiteralPath $OfflineMarker -PathType Leaf) -and (Test-Path -LiteralPath $WheelsRoot -PathType Container)) {
    & $PythonExe -m pip install --disable-pip-version-check --no-index --find-links $WheelsRoot -r (Join-Path $InstallRoot "requirements-runtime.txt")
} else {
    & $PythonExe -m pip install --disable-pip-version-check --upgrade "pip<27"
    if ($LASTEXITCODE -ne 0) { throw "pip could not be updated." }
    & $PythonExe -m pip install --disable-pip-version-check -r (Join-Path $InstallRoot "requirements-runtime.txt")
}
if ($LASTEXITCODE -ne 0) { throw "A required Python library failed to install." }

$PyWin32PostInstall = Join-Path $RuntimeRoot "Scripts\pywin32_postinstall.py"
if (Test-Path -LiteralPath $PyWin32PostInstall -PathType Leaf) {
    & $PythonExe $PyWin32PostInstall -install | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "pywin32 post-install returned a non-zero result. Outlook integration will be verified when the application opens."
    }
}

Write-Host "[4/6] Applying role and shared-data configuration..."
$ApiKeyValue = Get-EnvValue $ExistingEnv "GEMINI_API_KEY"
$NeedsConfiguration = ($InstallMode -eq "PROCESSOR") -and [string]::IsNullOrWhiteSpace($ApiKeyValue)

# Best-effort migration for rows processed by older versions.  Those rows may
# contain only the SMTP address even though the archived email.msg still has
# Outlook's real sender display name.  Recover it now so the Correos column can
# immediately show `Name <email>` after an upgrade.  Failure is non-fatal and
# the same migration is retried on every processor run.
if ($InstallMode -eq "PROCESSOR") {
    try {
        & $PythonExe (Join-Path $AppRoot "backfill_senders.py") | Out-Host
    } catch {
        Write-Warning "Existing sender names could not be enriched during installation. They will be retried automatically during email processing."
    }
}

Write-Host "[5/6] Creating shortcuts..."
$Shell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$ProductName"
New-Item -ItemType Directory -Force -Path $StartMenu | Out-Null

function New-AppShortcut([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkingDirectory, [string]$Description, [string]$IconLocation = "") {
    $Shortcut = $Shell.CreateShortcut($Path)
    $Shortcut.TargetPath = $Target
    $Shortcut.Arguments = $Arguments
    $Shortcut.WorkingDirectory = $WorkingDirectory
    $Shortcut.Description = $Description
    if (-not [string]::IsNullOrWhiteSpace($IconLocation)) { $Shortcut.IconLocation = "$IconLocation,0" }
    $Shortcut.Save()
}

$IconPath = Join-Path $AppRoot "assets\app.ico"
$DesktopLink = Join-Path $Desktop "$ProductName.lnk"
$StartMenuLink = Join-Path $StartMenu "$ProductName.lnk"
Remove-Item -LiteralPath $DesktopLink -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $StartMenuLink -Force -ErrorAction SilentlyContinue
$RoleDescription = if ($InstallMode -eq "PROCESSOR") { "$ProductName - Processing Computer" } else { "$ProductName - Viewer Computer" }
New-AppShortcut $DesktopLink $PythonWExe "`"$(Join-Path $AppRoot 'launch_desktop.pyw')`"" $AppRoot $RoleDescription $IconPath
New-AppShortcut $StartMenuLink $PythonWExe "`"$(Join-Path $AppRoot 'launch_desktop.pyw')`"" $AppRoot $RoleDescription $IconPath
New-AppShortcut (Join-Path $StartMenu "Configure $ProductName.lnk") "notepad.exe" "`"$ExistingEnv`"" $AppRoot "Open application configuration"
New-AppShortcut (Join-Path $StartMenu "Open shared data.lnk") "explorer.exe" "`"$SharedDataPath`"" $AppRoot "Open shared Email Assistant data"
New-AppShortcut (Join-Path $StartMenu "Open project archive.lnk") "explorer.exe" "`"$ArchiveRootPath`"" $AppRoot "Open engineering project archive"
New-AppShortcut (Join-Path $StartMenu "Open logs.lnk") "explorer.exe" "`"$(Join-Path $AppRoot 'logs')`"" $AppRoot "Open scheduler logs"
New-AppShortcut (Join-Path $StartMenu "Uninstall $ProductName.lnk") "powershell.exe" "-NoProfile -ExecutionPolicy Bypass -File `"$(Join-Path $MaintenanceRoot 'uninstall.ps1')`"" $InstallRoot "Uninstall application"
try {
    $IconRefresh = Join-Path $env:SystemRoot "System32\ie4uinit.exe"
    if (Test-Path -LiteralPath $IconRefresh) { Start-Process -FilePath $IconRefresh -ArgumentList "-show" -WindowStyle Hidden -Wait }
} catch { }

Write-Host "[6/6] Configuring automation for this role..."
$SchedulerInstalled = $false
if ($InstallMode -eq "PROCESSOR") {
    try {
        & (Join-Path $MaintenanceRoot "setup_scheduler.ps1")
        $SchedulerInstalled = $true
    } catch {
        Write-Warning "The application was installed, but automatic email processing could not be registered. $($_.Exception.Message)"
    }
} else {
    & (Join-Path $MaintenanceRoot "disable_all_schedulers.ps1")
}

if ($NeedsConfiguration) {
    Write-Host ""
    Write-Host "The processor configuration file will open now." -ForegroundColor Yellow
    Write-Host "Paste the Gemini API key after GEMINI_API_KEY=, save, and close Notepad."
    Start-Process notepad.exe -ArgumentList "`"$ExistingEnv`"" -Wait
}

Write-Host ""
Write-Host "$ProductName was installed successfully as $InstallMode." -ForegroundColor Green
Write-Host "Shared app data folder: $SharedDataPath"
Write-Host "Project archive root: $ArchiveRootPath"
Write-Host "Emails/attachments are filed under: $ArchiveRootPath\TRABAJOS <year>\<project/company>"
if ($InstallMode -eq "PROCESSOR") {
    Write-Host "Outlook mailbox: $TargetMailbox"
    if ($SchedulerInstalled) {
        Write-Host "Automatic processing runs every 30 minutes from 05:30 to 20:00 while this Windows user is signed in."
    } else {
        Write-Host "Automatic processing is not active yet. Use REPAIR_EMAIL_SCHEDULER.bat after checking Windows Task Scheduler." -ForegroundColor Yellow
    }
} else {
    Write-Host "Viewer mode: Outlook access, email processing, and scheduled tasks are disabled."
}
Write-Host ""
Start-Process -FilePath $PythonWExe -ArgumentList "`"$(Join-Path $AppRoot 'launch_desktop.pyw')`"" -WorkingDirectory $AppRoot
