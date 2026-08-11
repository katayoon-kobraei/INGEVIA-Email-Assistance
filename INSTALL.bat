@echo off
setlocal
cd /d "%~dp0"
title INGEVIA Email Assistant Installer

echo ======================================================
echo   INGEVIA Email AI Assistant - Portable Installer
echo ======================================================
echo.
echo Choose the role for this computer:
echo.
echo   1. Install as PROCESSING COMPUTER
echo      Connects to the configured boss Outlook mailbox,
echo      files emails/attachments into the engineering archive,
echo      writes shared app data, and runs scheduler.
echo.
echo   2. Install as VIEWER COMPUTER
echo      Reads the shared app data and engineering archive only.
echo      It does not connect to Outlook, process emails, or create tasks.
echo.
set /p choice=Enter 1 or 2: 
if "%choice%"=="1" goto processor
if "%choice%"=="2" goto viewer
echo.
echo Invalid selection.
pause
exit /b 1

:processor
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_installer\install.ps1" -PackageRoot "%~dp0." -InstallMode PROCESSOR
goto done

:viewer
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_installer\install.ps1" -PackageRoot "%~dp0." -InstallMode VIEWER
goto done

:done
if errorlevel 1 (
    echo.
    echo Installation failed. Read the error message above.
    echo.
    pause
    exit /b 1
)
echo.
echo Installation completed.
pause
