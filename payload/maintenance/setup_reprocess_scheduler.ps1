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
    throw "Scheduled reprocessing is disabled because this computer is installed in VIEWER mode."
}
$PythonW = Join-Path $InstallRoot "runtime\pythonw.exe"
$Launcher = Join-Path $AppRoot "launch_reprocess_unsorted.pyw"
$TaskName = "INGEVIA Email AI Assistant - Reprocess UNSORTED"

if (-not (Test-Path -LiteralPath $PythonW -PathType Leaf)) { throw "Python runtime not found: $PythonW" }
if (-not (Test-Path -LiteralPath $Launcher -PathType Leaf)) { throw "Reprocess UNSORTED launcher not found: $Launcher" }

# Same Windows Task Scheduler COM API approach as setup_scheduler.ps1 -- see
# its own comment for why (avoids a schtasks.exe quoting bug with the space
# in the task name).
#
# This task never touches Outlook -- reprocess_unsorted.py only reads
# already-saved files on disk and calls Gemini -- so, unlike the main email
# task, it isn't limited to business hours: it runs every 2 hours around
# the clock. Running when UNSORTED happens to be empty is harmless and
# cheap (the script just reports 0 entries and exits).
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $TaskDefinition = $Service.NewTask(0)

    $TaskDefinition.RegistrationInfo.Description = "Re-checks every email in this year's UNSORTED folder every 2 hours, without opening Outlook, and files or deletes what Gemini can now classify."
    $TaskDefinition.Settings.Enabled = $true
    $TaskDefinition.Settings.Hidden = $false
    $TaskDefinition.Settings.StartWhenAvailable = $true
    $TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $TaskDefinition.Settings.StopIfGoingOnBatteries = $false
    $TaskDefinition.Settings.MultipleInstances = 2  # TASK_INSTANCES_IGNORE_NEW
    $TaskDefinition.Settings.ExecutionTimeLimit = "PT1H"

    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $TaskDefinition.Principal.UserId = $CurrentUser
    $TaskDefinition.Principal.LogonType = 3  # TASK_LOGON_INTERACTIVE_TOKEN
    $TaskDefinition.Principal.RunLevel = 0   # TASK_RUNLEVEL_LUA (limited)

    $Action = $TaskDefinition.Actions.Create(0)  # TASK_ACTION_EXEC
    $Action.Path = $PythonW
    $Action.Arguments = ('"{0}"' -f $Launcher)
    $Action.WorkingDirectory = $AppRoot

    $Now = Get-Date
    for ($MinutesFromMidnight = 0; $MinutesFromMidnight -le 1320; $MinutesFromMidnight += 120) {
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
    throw "Windows Task Scheduler could not register the reprocessing task. Ensure the Task Scheduler service is enabled. Details: $($_.Exception.Message)"
}

Write-Host "Scheduler installed: $TaskName" -ForegroundColor Green