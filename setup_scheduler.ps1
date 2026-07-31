# setup_scheduler.ps1
$ErrorActionPreference = "Stop"

$TaskName = "Email AI Assistant"
$ProjectRoot = $PSScriptRoot
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\pythonw.exe"
$PipelineScript = Join-Path $ProjectRoot "src\pipeline.py"

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action = New-ScheduledTaskAction -Execute $PythonExe -Argument "`"$PipelineScript`"" -WorkingDirectory $ProjectRoot

# Runs every 30 minutes, but only between 05:30 and 20:00 -- starts at
# 05:30, repeats every 30 min for 14h30m (until 20:00), then goes
# quiet until 05:30 the next day. Running only during business hours
# roughly halves the number of Outlook connections opened per day
# compared to running 24/7, which also helps keep Outlook's resource
# usage down.
#
# The 05:30 run follows a ~9.5 hour overnight gap since the previous
# day's last run at 20:00 -- src/pipeline.py handles this itself by
# remembering the timestamp of its last successful run and looking
# back far enough to cover that whole gap automatically (see
# src/output/run_state.py), so nothing sent overnight gets missed.
# Nothing extra is needed here in the scheduler for that.
$Trigger = New-ScheduledTaskTrigger -Daily -At "05:30" `
    -RepetitionInterval (New-TimeSpan -Minutes 30) `
    -RepetitionDuration (New-TimeSpan -Hours 14 -Minutes 30)

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 25) `
    -MultipleInstances IgnoreNew `
    -Compatibility Win8

$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Settings $Settings -Principal $Principal `
    -Description "Checks Outlook for new emails, classifies them, and saves them to folders. Runs every 30 min, 05:30-20:00 daily."

Write-Host "Done. '$TaskName' will now run every 30 minutes between 05:30 and 20:00, every day."