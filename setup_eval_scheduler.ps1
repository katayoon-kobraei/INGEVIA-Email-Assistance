$ErrorActionPreference = "Stop"

$TaskName = "Email AI Assistant - Daily Eval"
$ProjectRoot = $PSScriptRoot
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\pythonw.exe"
$EvalScript = Join-Path $ProjectRoot "eval_project_agent.py"

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$EvalScript`"" -WorkingDirectory $ProjectRoot

# Every morning at 8:00 AM -- change -At to taste.
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

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