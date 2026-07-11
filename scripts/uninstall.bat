@echo off
REM DeepExcel Uninstaller - double-click to run
cd /d "%~dp0"

set "PS_EXE=powershell.exe"
where %PS_EXE% >nul 2>nul
if errorlevel 1 set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1" -Unregister
pause
