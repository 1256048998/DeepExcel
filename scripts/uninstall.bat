@echo off
REM DeepExcel Uninstaller - double-click to run
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1" -Unregister
pause
