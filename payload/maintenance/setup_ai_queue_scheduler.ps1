$ErrorActionPreference = "Stop"
$InstallRoot = Split-Path -Parent $PSScriptRoot
$AppRoot = Join-Path $InstallRoot "app"
$EnvPath = Join-Path $AppRoot ".env"
$Mode = "PROCESSOR"
if (Test-Path -LiteralPath $EnvPath -PathType Leaf) {
    $ModeLine = Get-Content -LiteralPath $EnvPath -ErrorAction SilentlyContinue |
        Where-Object { $_ -match '^\s*APP_MODE\s*=' } | Select-Object -First 1
    if ($ModeLine) { $Mode = (($ModeLine -split '=', 2)[1]).Trim().ToUpperInvariant() }
}
if ($Mode -eq "VIEWER") {
    throw "AI queue scheduling is disabled because this computer is installed in VIEWER mode."
}

$PythonW = Join-Path $InstallRoot "runtime\pythonw.exe"
$Launcher = Join-Path $AppRoot "launch_ai_queue.pyw"
$TaskName = "INGEVIA Email AI Assistant - AI Queue"
if (-not (Test-Path -LiteralPath $PythonW -PathType Leaf)) { throw "Python runtime not found: $PythonW" }
if (-not (Test-Path -LiteralPath $Launcher -PathType Leaf)) { throw "AI queue launcher not found: $Launcher" }

$Service = New-Object -ComObject "Schedule.Service"
$Service.Connect()
$RootFolder = $Service.GetFolder("\")
$TaskDefinition = $Service.NewTask(0)
$TaskDefinition.RegistrationInfo.Description = "Processes up to 2 locally captured emails with Gemini every 2 minutes without opening Outlook; v1.26.9 also repairs recoverable v1.25 path failures."
$TaskDefinition.Settings.Enabled = $true
$TaskDefinition.Settings.Hidden = $false
$TaskDefinition.Settings.StartWhenAvailable = $true
$TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
$TaskDefinition.Settings.StopIfGoingOnBatteries = $false
$TaskDefinition.Settings.MultipleInstances = 2  # TASK_INSTANCES_IGNORE_NEW
$TaskDefinition.Settings.ExecutionTimeLimit = "PT20M"

$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$TaskDefinition.Principal.UserId = $CurrentUser
$TaskDefinition.Principal.LogonType = 3
$TaskDefinition.Principal.RunLevel = 0

$Action = $TaskDefinition.Actions.Create(0)
$Action.Path = $PythonW
$Action.Arguments = ('"{0}"' -f $Launcher)
$Action.WorkingDirectory = $AppRoot

# Daily 05:30 -> 20:00 window, repeating every 2 minutes.  One repeating
# trigger is much lighter than registering hundreds of individual triggers.
$Now = Get-Date
$TodayStart = [DateTime]::Today.AddHours(5).AddMinutes(30)
$TodayEnd = [DateTime]::Today.AddHours(20)

# If installation happens during the working window, add a one-off trigger so
# the queue starts within about one minute instead of waiting until tomorrow.
# Before 05:30 the normal daily trigger below already starts today, so do not
# create a duplicate repeating trigger for the same window.
if ($Now -ge $TodayStart -and $Now -lt $TodayEnd) {
    $StartToday = $Now.AddMinutes(1)
    $Remaining = $TodayEnd - $StartToday
    if ($Remaining.TotalMinutes -ge 2) {
        $Immediate = $TaskDefinition.Triggers.Create(1) # TASK_TRIGGER_TIME
        $Immediate.StartBoundary = $StartToday.ToString("yyyy-MM-dd'T'HH:mm:ss")
        $Immediate.Enabled = $true
        $Immediate.Repetition.Interval = "PT2M"
        $Immediate.Repetition.Duration = ("PT{0}M" -f [Math]::Max(2, [Math]::Floor($Remaining.TotalMinutes)))
        $Immediate.Repetition.StopAtDurationEnd = $true
    }
}

$DailyStart = if ($Now -lt $TodayStart) { $TodayStart } else { $TodayStart.AddDays(1) }
$Daily = $TaskDefinition.Triggers.Create(2) # TASK_TRIGGER_DAILY
$Daily.StartBoundary = $DailyStart.ToString("yyyy-MM-dd'T'HH:mm:ss")
$Daily.DaysInterval = 1
$Daily.Enabled = $true
$Daily.Repetition.Interval = "PT2M"
$Daily.Repetition.Duration = "PT14H30M"
$Daily.Repetition.StopAtDurationEnd = $true

$null = $RootFolder.RegisterTaskDefinition($TaskName, $TaskDefinition, 6, $CurrentUser, $null, 3, $null)
Write-Host "AI queue scheduler installed: $TaskName (every 2 minutes, max 2 staged emails per run)" -ForegroundColor Green
