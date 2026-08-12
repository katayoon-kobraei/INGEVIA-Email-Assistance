$ErrorActionPreference = "SilentlyContinue"
$TaskName = "INGEVIA Email AI Assistant - Reprocess UNSORTED"
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    $RootFolder.DeleteTask($TaskName, 0)
} catch { }
Write-Host "Reprocess UNSORTED scheduler removed (or it was already absent)." -ForegroundColor Green