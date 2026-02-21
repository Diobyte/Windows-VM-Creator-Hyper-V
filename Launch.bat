@echo off
title Hyper-V Toolkit Launcher - Version 1
color 0A
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%HyperV-Toolkit.ps1"
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "LAUNCH_LOG=%TEMP%\HyperV-Toolkit-Launcher.log"

call :log "Launcher started"
call :log "Script path: %SCRIPT_PATH%"

:: Prefer 64-bit Windows PowerShell even when launched from a 32-bit host
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if defined PROCESSOR_ARCHITEW6432 (
    if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
        set "POWERSHELL_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
        call :log "Using Sysnative 64-bit PowerShell: %POWERSHELL_EXE%"
    )
)

if not exist "%POWERSHELL_EXE%" (
    set "POWERSHELL_EXE=PowerShell.exe"
    call :log "Fallback to PATH PowerShell.exe"
)

call :log "PowerShell executable: %POWERSHELL_EXE%"

if not exist "%SCRIPT_PATH%" (
    call :log "ERROR: script not found"
    echo.
    echo  ERROR: Could not find "%SCRIPT_PATH%"
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)

:: Check for admin privileges using WindowsPrincipal (works even if Server service is disabled)
"%POWERSHELL_EXE%" -NoProfile -Command "$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent(); if($p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){exit 0}else{exit 1}" >nul 2>&1
if %errorLevel% neq 0 (
    call :log "Not elevated; requesting RunAs"
    if /i "%~1"=="--elevated" (
        call :log "ERROR: elevation failed or was declined"
        echo.
        echo  ERROR: UAC elevation failed or was declined. Please run as Administrator.
        echo  Press any key to exit...
        pause >nul
        exit /b 1
    )
    echo.
    echo  Requesting administrator privileges...
    echo.
    set "LAUNCHER_PATH=%~f0"
    "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$launcher = $env:LAUNCHER_PATH; Start-Process -FilePath $launcher -Verb RunAs -ArgumentList @('--elevated')"
    call :log "Elevation request dispatched"
    exit /b
)

if /i "%~1"=="--elevated" (
    title Hyper-V Toolkit Launcher - Version 1 [Administrator]
    call :log "Running in elevated launcher instance"
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

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Sta -File "%SCRIPT_PATH%"

set "LAUNCH_EXIT_CODE=%errorLevel%"
call :log "Toolkit exited with code: %LAUNCH_EXIT_CODE%"

if %LAUNCH_EXIT_CODE% neq 0 (
    echo.
    echo  Launcher detected a startup error. Press any key to exit...
    pause >nul
)

endlocal
exit /b %LAUNCH_EXIT_CODE%

:log
echo [%date% %time%] %~1>>"%LAUNCH_LOG%"
goto :eof
