@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\setup_scheduler.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\maintenance\setup_ai_queue_scheduler.ps1"
echo.
pause
