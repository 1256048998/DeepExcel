@echo off
REM DeepExcel Installer - double-click to run
REM Uses Windows PowerShell with full path fallback in case PATH is broken.

cd /d "%~dp0"

REM Try powershell.exe from PATH first, then fall back to full System32 path.
set "PS_EXE=powershell.exe"
where %PS_EXE% >nul 2>nul
if errorlevel 1 set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS_EXE%" (
    echo ERROR: Windows PowerShell not found on this system.
    echo DeepExcel requires Windows PowerShell (built into Windows 7+).
    pause
    exit /b 1
)

echo Launching DeepExcel installer via PowerShell...
echo.

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1"

echo.
pause
