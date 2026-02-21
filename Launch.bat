@echo off
title Hyper-V Toolkit Launcher - Version 1
color 0A
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%HyperV-Toolkit.ps1"

if not exist "%SCRIPT_PATH%" (
    echo.
    echo  ERROR: Could not find "%SCRIPT_PATH%"
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)

:: Check for admin privileges (net session is more reliable than fltmc)
net session >nul 2>&1
if %errorLevel% neq 0 (
    if /i "%~1"=="--elevated" (
        echo.
        echo  ERROR: UAC elevation failed or was declined. Please run as Administrator.
        echo  Press any key to exit...
        pause >nul
        exit /b 1
    )
    echo.
    echo  Requesting administrator privileges...
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs -ArgumentList '--elevated'"
    exit /b
)

if /i "%~1"=="--elevated" shift

:: Change to script directory
cd /d "%SCRIPT_DIR%"

echo.
echo  =============================================
echo   Hyper-V Toolkit ^| Version 1 ^| Diobyte ^| Made with love
echo  =============================================
echo.
echo  Launching toolkit...
echo.

PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%SCRIPT_PATH%"

if %errorLevel% neq 0 (
    echo.
    echo  Launcher detected a startup error. Press any key to exit...
    pause >nul
)

endlocal
