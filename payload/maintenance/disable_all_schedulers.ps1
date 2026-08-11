$ErrorActionPreference = "SilentlyContinue"
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    @("INGEVIA Email AI Assistant", "INGEVIA Email AI Assistant - Daily Evaluation") | ForEach-Object {
        try { $RootFolder.DeleteTask($_, 0) } catch { }
    }
} catch { }
Write-Host "Email AI Assistant scheduled tasks removed (or they were already absent)." -ForegroundColor Green
