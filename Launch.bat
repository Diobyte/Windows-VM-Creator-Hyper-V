@echo off
title Hyper-V Toolkit Launcher - Version 1
color 0A
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%HyperV-Toolkit.ps1"
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%POWERSHELL_EXE%" (
    set "POWERSHELL_EXE=PowerShell.exe"
)

if not exist "%SCRIPT_PATH%" (
    echo.
    echo  ERROR: Could not find "%SCRIPT_PATH%"
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)

:: Check for admin privileges using WindowsPrincipal (works even if Server service is disabled)
"%POWERSHELL_EXE%" -NoProfile -Command "$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); if($p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){exit 0}else{exit 1}" >nul 2>&1
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
    "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy RemoteSigned -Command "Start-Process -FilePath '%~f0' -Verb RunAs -ArgumentList @('--elevated')"
    exit /b
)

if /i "%~1"=="--elevated" (
    title Hyper-V Toolkit Launcher - Version 1 [Administrator]
    shift
)

:: Change to script directory
cd /d "%SCRIPT_DIR%"

echo.
echo  =============================================
echo   Hyper-V Toolkit ^| Version 1 ^| Diobyte ^| Made with love
echo  =============================================
echo.
echo  Script: %SCRIPT_PATH%
echo  Launching toolkit...
echo.

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy RemoteSigned -Sta -File "%SCRIPT_PATH%"

set "LAUNCH_EXIT_CODE=%errorLevel%"

if %LAUNCH_EXIT_CODE% neq 0 (
    echo.
    echo  Launcher detected a startup error. Press any key to exit...
    pause >nul
)

endlocal
exit /b %LAUNCH_EXIT_CODE%
