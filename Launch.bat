@echo off
title Hyper-V Toolkit Launcher - Version 1
color 0A

:: Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo  Requesting administrator privileges...
    echo.
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Change to script directory
cd /d "%~dp0"

echo.
echo  =============================================
echo   Hyper-V Toolkit ^| Version 1 ^| Diobyte ^| Made with love
echo  =============================================
echo.
echo  Launching toolkit...
echo.

PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0HyperV-Toolkit.ps1"

if %errorLevel% neq 0 (
    echo.
    echo  An error occurred. Press any key to exit...
    pause >nul
)
