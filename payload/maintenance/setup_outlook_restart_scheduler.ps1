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
    throw "Outlook restart scheduling is disabled because this computer is installed in VIEWER mode."
}

$PythonW = Join-Path $InstallRoot "runtime\pythonw.exe"
$Launcher = Join-Path $AppRoot "launch_outlook_restart.pyw"
$TaskName = "INGEVIA Email AI Assistant - Outlook Daily Restart"
if (-not (Test-Path -LiteralPath $PythonW -PathType Leaf)) { throw "Python runtime not found: $PythonW" }
if (-not (Test-Path -LiteralPath $Launcher -PathType Leaf)) { throw "Outlook restart launcher not found: $Launcher" }

$Service = New-Object -ComObject "Schedule.Service"
$Service.Connect()
$RootFolder = $Service.GetFolder("\")
$TaskDefinition = $Service.NewTask(0)
$TaskDefinition.RegistrationInfo.Description = (
    "Gracefully closes and relaunches OUTLOOK.EXE once nightly (with a second " +
    "attempt two hours later if the first could not close it) so its COM/MAPI " +
    "resource usage resets on a schedule, without requiring the PC itself to be " +
    "restarted. Never force-kills Outlook -- if it will not close gracefully, " +
    "this leaves it running and the next scheduled attempt tries again."
)
$TaskDefinition.Settings.Enabled = $true
$TaskDefinition.Settings.Hidden = $false
$TaskDefinition.Settings.StartWhenAvailable = $true
$TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
$TaskDefinition.Settings.StopIfGoingOnBatteries = $false
$TaskDefinition.Settings.MultipleInstances = 2  # TASK_INSTANCES_IGNORE_NEW
$TaskDefinition.Settings.ExecutionTimeLimit = "PT10M"

$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$TaskDefinition.Principal.UserId = $CurrentUser
$TaskDefinition.Principal.LogonType = 3
$TaskDefinition.Principal.RunLevel = 0

$Action = $TaskDefinition.Actions.Create(0)
$Action.Path = $PythonW
$Action.Arguments = ('"{0}"' -f $Launcher)
$Action.WorkingDirectory = $AppRoot

# Two fixed daily triggers rather than one -- both well outside the 05:30-20:00
# capture window, so this never competes with live email processing:
#   21:00 -- first (and normally only) attempt.
#   23:00 -- retry ONLY needed if 21:00 found Outlook stuck and would not close
#            gracefully. The 23:00 run is the exact same idempotent script; if
#            21:00 already succeeded, 23:00 just confirms Outlook is running
#            and does nothing further.
$Today = Get-Date -Hour 0 -Minute 0 -Second 0
$FirstRun = $Today.AddHours(21)
$SecondRun = $Today.AddHours(23)
if ($FirstRun -lt (Get-Date)) { $FirstRun = $FirstRun.AddDays(1) }
if ($SecondRun -lt (Get-Date)) { $SecondRun = $SecondRun.AddDays(1) }

$Trigger1 = $TaskDefinition.Triggers.Create(2)  # TASK_TRIGGER_DAILY
$Trigger1.StartBoundary = $FirstRun.ToString("yyyy-MM-dd'T'HH:mm:ss")
$Trigger1.DaysInterval = 1
$Trigger1.Enabled = $true

$Trigger2 = $TaskDefinition.Triggers.Create(2)  # TASK_TRIGGER_DAILY
$Trigger2.StartBoundary = $SecondRun.ToString("yyyy-MM-dd'T'HH:mm:ss")
$Trigger2.DaysInterval = 1
$Trigger2.Enabled = $true

$null = $RootFolder.RegisterTaskDefinition($TaskName, $TaskDefinition, 6, $CurrentUser, $null, 3, $null)
Write-Host "Outlook daily restart scheduler installed: $TaskName (21:00, retry 23:00 if needed)" -ForegroundColor Green