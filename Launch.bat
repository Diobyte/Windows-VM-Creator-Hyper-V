@echo off
title Hyper-V Toolkit Launcher
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%HyperV-Toolkit.ps1"
set "POWERSHELL_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "LAUNCH_LOG=%TEMP%\HyperV-Toolkit-Launcher.log"

:: Rotate launcher log if larger than 100 KB to prevent unbounded growth
if exist "%LAUNCH_LOG%" (
    for %%F in ("%LAUNCH_LOG%") do if %%~zF gtr 102400 (
        del /f /q "%LAUNCH_LOG%" >nul 2>&1
    )
)

call :log "Launcher started"
call :log "Script path: %SCRIPT_PATH%"

:: ----------------------------------------------------------------
:: Prefer 64-bit Windows PowerShell even when launched from 32-bit host
:: ----------------------------------------------------------------
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if defined PROCESSOR_ARCHITEW6432 (
    if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
        set "POWERSHELL_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
        call :log "Using Sysnative 64-bit PowerShell"
    )
)

if not exist "%POWERSHELL_EXE%" (
    set "POWERSHELL_EXE=PowerShell.exe"
    call :log "Fallback: using PATH PowerShell.exe"
)

call :log "PowerShell executable: %POWERSHELL_EXE%"

:: ----------------------------------------------------------------
:: Script existence check
:: ----------------------------------------------------------------
if not exist "%SCRIPT_PATH%" (
    call :log "ERROR: script not found at: %SCRIPT_PATH%"
    echo.
    echo  ERROR: Could not find:
    echo         %SCRIPT_PATH%
    echo.
    echo  Press any key to exit...
    pause >nul
    exit /b 1
)

:: ----------------------------------------------------------------
:: Admin check (uses WindowsPrincipal — works without Server service)
:: ----------------------------------------------------------------
"%POWERSHELL_EXE%" -NoProfile -Command ^
    "$p=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent();" ^
    "if($p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){exit 0}else{exit 1}" >nul 2>&1

if %errorLevel% neq 0 (
    call :log "Not elevated — requesting RunAs"
    if /i "%~1"=="--elevated" (
        call :log "ERROR: elevation failed or was declined"
        echo.
        echo  ERROR: UAC elevation failed or was declined.
        echo  Please right-click Launch.bat and choose "Run as administrator".
        echo.
        echo  Press any key to exit...
        pause >nul
        exit /b 1
    )
    echo.
    echo  Requesting administrator privileges...
    echo.
    set "LAUNCHER_PATH=%~f0"
    "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Command ^
        "try { Start-Process -FilePath $env:LAUNCHER_PATH -Verb RunAs" ^
        "      -ArgumentList '--elevated' -ErrorAction Stop; exit 0" ^
        "} catch { exit 1 }" >nul 2>&1
    if %errorLevel% neq 0 (
        call :log "ERROR: elevation request failed"
        echo.
        echo  ERROR: Could not request elevation.
        echo  Please right-click Launch.bat and choose "Run as administrator".
        echo.
        echo  Press any key to exit...
        pause >nul
        exit /b 1
    )
    call :log "Elevation request dispatched"
    exit /b
)

:: ----------------------------------------------------------------
:: Elevated instance setup
:: ----------------------------------------------------------------
if /i "%~1"=="--elevated" (
    title Hyper-V Toolkit Launcher  [Administrator]
    call :log "Running elevated launcher instance"
    shift
)

cd /d "%SCRIPT_DIR%"

:: ----------------------------------------------------------------
:: Banner
:: ----------------------------------------------------------------
color 0A
echo.
echo  =============================================
echo   Hyper-V VM Creator  ^|  Diobyte
echo  =============================================
echo.
echo  Launching... the console will close automatically.
echo.

:: ----------------------------------------------------------------
:: Launch the GUI — invoke PowerShell directly (inheriting this console).
:: PowerShell processes -WindowStyle Hidden by calling ShowWindow(SW_HIDE)
:: on the inherited console handle, hiding it after the brief banner flash.
:: Using "start /wait" would spawn a NEW console that stays visible the
:: whole time the WinForms GUI runs, so we invoke PowerShell directly.
:: ----------------------------------------------------------------
call :log "Launching toolkit (hidden console)"
"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden ^
    -File "%SCRIPT_PATH%"

set "LAUNCH_EXIT_CODE=%errorLevel%"
call :log "Toolkit exited with code: %LAUNCH_EXIT_CODE%"

if %LAUNCH_EXIT_CODE% neq 0 (
    color 0C
    echo.
    echo  The toolkit exited with an error (code: %LAUNCH_EXIT_CODE%).
    echo  Check the log: %LAUNCH_LOG%
    echo.
    echo  Press any key to exit...
    pause >nul
)

endlocal
exit /b %LAUNCH_EXIT_CODE%

:: ----------------------------------------------------------------
:log
echo [%date% %time%] %~1>>"%LAUNCH_LOG%"
goto :eof
