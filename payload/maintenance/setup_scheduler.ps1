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
    throw "Scheduled processing is disabled because this computer is installed in VIEWER mode."
}
$PythonW = Join-Path $InstallRoot "runtime\pythonw.exe"
$Launcher = Join-Path $AppRoot "launch_pipeline.pyw"
$TaskName = "INGEVIA Email AI Assistant"

if (-not (Test-Path -LiteralPath $PythonW -PathType Leaf)) { throw "Python runtime not found: $PythonW" }
if (-not (Test-Path -LiteralPath $Launcher -PathType Leaf)) { throw "Pipeline launcher not found: $Launcher" }

# Use the Windows Task Scheduler COM API rather than schtasks.exe. This keeps
# task names and paths containing spaces as single values and avoids the native
# command-line quoting problem that caused 'Invalid argument/option - Email'.
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $TaskDefinition = $Service.NewTask(0)

    $TaskDefinition.RegistrationInfo.Description = "Processes INGEVIA Outlook email every 30 minutes between 05:30 and 20:00."
    $TaskDefinition.Settings.Enabled = $true
    $TaskDefinition.Settings.Hidden = $false
    $TaskDefinition.Settings.StartWhenAvailable = $true
    $TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $TaskDefinition.Settings.StopIfGoingOnBatteries = $false
    $TaskDefinition.Settings.MultipleInstances = 2  # TASK_INSTANCES_IGNORE_NEW
    $TaskDefinition.Settings.ExecutionTimeLimit = "PT2H"

    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $TaskDefinition.Principal.UserId = $CurrentUser
    $TaskDefinition.Principal.LogonType = 3  # TASK_LOGON_INTERACTIVE_TOKEN
    $TaskDefinition.Principal.RunLevel = 0   # TASK_RUNLEVEL_LUA (limited)

    $Action = $TaskDefinition.Actions.Create(0)  # TASK_ACTION_EXEC
    $Action.Path = $PythonW
    $Action.Arguments = ('"{0}"' -f $Launcher)
    $Action.WorkingDirectory = $AppRoot

    $Now = Get-Date
    for ($MinutesFromMidnight = 330; $MinutesFromMidnight -le 1200; $MinutesFromMidnight += 30) {
        $Hour = [Math]::Floor($MinutesFromMidnight / 60)
        $Minute = $MinutesFromMidnight % 60
        $FirstRun = [DateTime]::Today.AddHours($Hour).AddMinutes($Minute)
        if ($FirstRun -le $Now) { $FirstRun = $FirstRun.AddDays(1) }

        $Trigger = $TaskDefinition.Triggers.Create(2)  # TASK_TRIGGER_DAILY
        $Trigger.StartBoundary = $FirstRun.ToString("yyyy-MM-dd'T'HH:mm:ss")
        $Trigger.DaysInterval = 1
        $Trigger.Enabled = $true
    }

    # TASK_CREATE_OR_UPDATE = 6. No password is needed because the task runs only
    # while the same Windows user is signed in.
    $null = $RootFolder.RegisterTaskDefinition($TaskName, $TaskDefinition, 6, $CurrentUser, $null, 3, $null)
} catch {
    throw "Windows Task Scheduler could not register the email task. Ensure the Task Scheduler service is enabled. Details: $($_.Exception.Message)"
}

Write-Host "Scheduler installed: $TaskName" -ForegroundColor Green
