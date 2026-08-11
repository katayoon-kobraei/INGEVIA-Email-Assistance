$ErrorActionPreference = "SilentlyContinue"
$TaskName = "INGEVIA Email AI Assistant - Daily Evaluation"
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $RootFolder.DeleteTask($TaskName, 0)
} catch { }
Write-Host "Daily evaluation scheduler removed (or it was already absent)." -ForegroundColor Green
