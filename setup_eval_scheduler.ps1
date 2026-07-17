# setup_eval_scheduler.ps1
#
# One-time setup, run on YOUR OWN dev machine only -- do NOT include
# this in the boss's production install. Registers a WEEKLY Task
# Scheduler job that re-runs eval_project_agent.py against the real
# archive and refreshes eval_results.xlsx automatically, so the
# latest accuracy numbers are already there whenever you check in.
#
# This is separate from setup_scheduler.ps1 (the production email
# pipeline task) -- different task name, won't conflict with it.

$ErrorActionPreference = "Stop"

$TaskName = "Email AI Assistant - Daily Eval"
$ProjectRoot = $PSScriptRoot
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\pythonw.exe"
$EvalScript = Join-Path $ProjectRoot "eval_project_agent.py"

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$EvalScript`"" -WorkingDirectory $ProjectRoot

# Every morning at 8:00 AM -- change -At to taste.
$Trigger = New-ScheduledTaskTrigger -Daily -At 8am

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -MultipleInstances IgnoreNew `
    -Compatibility Win8

$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Settings $Settings -Principal $Principal `
    -Description "Daily: re-runs the classifier against the real archive and refreshes eval_results.xlsx for review."

Write-Host "Done. '$TaskName' will now run every morning at 8:00 AM."
Write-Host "Check eval_results.xlsx afterward -- it's overwritten each run, so pull out anything you want to keep before tomorrow's run replaces it."