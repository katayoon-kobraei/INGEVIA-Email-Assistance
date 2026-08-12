$ErrorActionPreference = "SilentlyContinue"
$InstallRoot = Split-Path -Parent $PSScriptRoot
try {
    $Service = New-Object -ComObject "Schedule.Service"
    $Service.Connect()
    $RootFolder = $Service.GetFolder("\")
    @("INGEVIA Email AI Assistant", "INGEVIA Email AI Assistant - Daily Evaluation", "INGEVIA Email AI Assistant - Reprocess UNSORTED") | ForEach-Object {
        try { $RootFolder.DeleteTask($_, 0) } catch { }
    }
} catch { }
$Desktop = [Environment]::GetFolderPath("Desktop")
$StartMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\INGEVIA Email Assistant"
Remove-Item (Join-Path $Desktop "INGEVIA Email Assistant.lnk") -Force -ErrorAction SilentlyContinue
Remove-Item $StartMenu -Recurse -Force -ErrorAction SilentlyContinue
Start-Process cmd.exe -WindowStyle Hidden -ArgumentList "/c timeout /t 2 /nobreak >nul & rmdir /s /q `"$InstallRoot`""
Write-Host "INGEVIA Email Assistant has been uninstalled. The email archive/output folder was not deleted." -ForegroundColor Green
