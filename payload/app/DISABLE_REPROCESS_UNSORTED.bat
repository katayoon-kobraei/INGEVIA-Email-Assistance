@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\disable_reprocess_scheduler.ps1"
echo.
pause