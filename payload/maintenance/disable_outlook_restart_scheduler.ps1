$ErrorActionPreference = "SilentlyContinue"
$TaskName = "INGEVIA Email AI Assistant - Outlook Daily Restart"
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $RootFolder.DeleteTask($TaskName, 0)
} catch { }
Write-Host "Outlook daily restart scheduler removed (or it was already absent)." -ForegroundColor Green