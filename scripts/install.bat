@echo off
cd /d "%~dp0"
echo ==========================================
echo   DeepExcel Installer
echo ==========================================
echo.

REM Unblock the PowerShell script first (in case it has MOTW from download)
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command "Unblock-File -Path '%~dp0register-user.ps1' -ErrorAction SilentlyContinue; Unblock-File -Path '%~dp0diagnose.ps1' -ErrorAction SilentlyContinue"

REM Run the registration script
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1"

echo.
echo ==========================================
pause
