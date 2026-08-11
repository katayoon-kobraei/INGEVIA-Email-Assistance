param([Parameter(Mandatory=$true)][string]$PackageRoot)
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Normalize-InstallerPath([string]$Path) {
    $Path = $Path.Trim().Trim('"').Trim("'")
    try { return [System.IO.Path]::GetFullPath($Path) }
    catch { throw "The installer folder path is invalid: $Path" }
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
    Invoke-WebRequest -Uri $Uri -OutFile $Destination -UseBasicParsing
    if (-not (Get-ValidCachedFile $Destination $MinimumBytes)) {
        Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
        throw "The downloaded $Description file is incomplete or invalid."
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
            if (-not $HasImportSite) { $Updated += 'import site'; $HasImportSite = $true }
            continue
        }
        if ($Trimmed -ieq 'Lib\site-packages' -or $Trimmed -ieq 'Lib/site-packages') {
            if (-not $HasSitePackages) { $Updated += 'Lib\site-packages'; $HasSitePackages = $true }
            continue
        }
        $Updated += $Line
    }
    if (-not $HasSitePackages) { $Updated += 'Lib\site-packages' }
    if (-not $HasImportSite) { $Updated += 'import site' }
    Set-Content -LiteralPath $PthFile.FullName -Value $Updated -Encoding ASCII
    New-Item -ItemType Directory -Force -Path (Join-Path $Root 'Lib\site-packages') | Out-Null
}

$PackageRoot = Normalize-InstallerPath $PackageRoot
if (-not (Test-Path -LiteralPath $PackageRoot -PathType Container)) { throw "The installer folder could not be found: $PackageRoot" }
$PackageRoot = (Get-Item -LiteralPath $PackageRoot).FullName
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$CacheRoot = Join-Path $PackageRoot "_offline_cache"
$WheelsRoot = Join-Path $CacheRoot "wheels"
$BuilderRoot = Join-Path $CacheRoot "builder_python"
$PythonArchive = Join-Path $CacheRoot "python-3.11.9-embed-amd64.zip"
$PythonUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"
$GetPipScript = Join-Path $CacheRoot "get-pip.py"
$GetPipUrl = "https://bootstrap.pypa.io/get-pip.py"
$Requirements = Join-Path $PackageRoot "payload\requirements-runtime.txt"
$Marker = Join-Path $CacheRoot "CACHE_COMPLETE.txt"

New-Item -ItemType Directory -Force -Path $CacheRoot, $WheelsRoot | Out-Null
Remove-Item -LiteralPath $Marker -Force -ErrorAction SilentlyContinue

Write-Host "Preparing a fully offline copy of the installer..." -ForegroundColor Cyan
Write-Host "[1/3] Caching the official portable Python runtime..."
Download-RequiredFile $PythonUrl $PythonArchive 8000000 "official portable Python runtime"
Download-RequiredFile $GetPipUrl $GetPipScript 500000 "official pip bootstrap script"

$BuilderPython = Join-Path $BuilderRoot "python.exe"
if (-not (Test-Path -LiteralPath $BuilderPython -PathType Leaf)) {
    Write-Host "[2/3] Creating a temporary package-building Python..."
    Remove-Item -LiteralPath $BuilderRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path $BuilderRoot | Out-Null
    Expand-Archive -LiteralPath $PythonArchive -DestinationPath $BuilderRoot -Force
    Enable-EmbeddedPythonSite $BuilderRoot
    $BootstrapArguments = @($GetPipScript, '--disable-pip-version-check', 'pip<27')
    & $BuilderPython @BootstrapArguments
    if ($LASTEXITCODE -ne 0) { throw "Temporary Python pip setup failed." }
} else {
    Write-Host "[2/3] Temporary package-building Python already present."
}

Write-Host "[3/3] Downloading all Windows dependency wheels..."
Remove-Item (Join-Path $WheelsRoot "*") -Recurse -Force -ErrorAction SilentlyContinue
& $BuilderPython -m pip install --disable-pip-version-check --upgrade "pip<27"
if ($LASTEXITCODE -ne 0) { throw "Temporary pip update failed." }
& $BuilderPython -m pip download --disable-pip-version-check --only-binary=:all: --dest $WheelsRoot -r $Requirements
if ($LASTEXITCODE -ne 0) { throw "One or more offline dependency files could not be downloaded." }
"Offline cache completed: $(Get-Date -Format s)" | Set-Content -LiteralPath $Marker -Encoding UTF8

Write-Host ""
Write-Host "Offline cache completed successfully." -ForegroundColor Green
Write-Host "You can now copy or ZIP this entire installer folder and install on other 64-bit Windows PCs without internet."
