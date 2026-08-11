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
$Launcher = Join-Path $AppRoot "launch_eval.pyw"
$TaskName = "INGEVIA Email AI Assistant - Daily Evaluation"

if (-not (Test-Path -LiteralPath $PythonW -PathType Leaf)) { throw "Python runtime not found: $PythonW" }
if (-not (Test-Path -LiteralPath $Launcher -PathType Leaf)) { throw "Evaluation launcher not found: $Launcher" }

try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $TaskDefinition = $Service.NewTask(0)

    $TaskDefinition.RegistrationInfo.Description = "Runs the optional INGEVIA Email AI Assistant evaluation once per day at 08:00."
    $TaskDefinition.Settings.Enabled = $true
    $TaskDefinition.Settings.Hidden = $false
    $TaskDefinition.Settings.StartWhenAvailable = $true
    $TaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $TaskDefinition.Settings.StopIfGoingOnBatteries = $false
    $TaskDefinition.Settings.MultipleInstances = 2
    $TaskDefinition.Settings.ExecutionTimeLimit = "PT4H"

    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $TaskDefinition.Principal.UserId = $CurrentUser
    $TaskDefinition.Principal.LogonType = 3
    $TaskDefinition.Principal.RunLevel = 0

    $Action = $TaskDefinition.Actions.Create(0)
    $Action.Path = $PythonW
    $Action.Arguments = ('"{0}"' -f $Launcher)
    $Action.WorkingDirectory = $AppRoot

    $FirstRun = [DateTime]::Today.AddHours(8)
    if ($FirstRun -le (Get-Date)) { $FirstRun = $FirstRun.AddDays(1) }
    $Trigger = $TaskDefinition.Triggers.Create(2)
    $Trigger.StartBoundary = $FirstRun.ToString("yyyy-MM-dd'T'HH:mm:ss")
    $Trigger.DaysInterval = 1
    $Trigger.Enabled = $true

    $null = $RootFolder.RegisterTaskDefinition($TaskName, $TaskDefinition, 6, $CurrentUser, $null, 3, $null)
} catch {
    throw "Windows Task Scheduler could not register the daily evaluation task. Ensure the Task Scheduler service is enabled. Details: $($_.Exception.Message)"
}

Write-Host "Daily evaluation scheduler installed: $TaskName" -ForegroundColor Green
