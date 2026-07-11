@echo off
cd /d "%~dp0"
echo ==========================================
echo   DeepExcel Installer
echo ==========================================
echo.

REM Unblock all files in this folder (removes Mark of the Web from download)
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -Command "Get-ChildItem -Path '%~dp0' -Recurse -File | ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }"
echo Files unblocked.
echo.

REM Try the .exe installer first (full install with HKLM + clean path)
if exist "%~dp0DeepExcelInstaller.exe" (
    echo Launching DeepExcelInstaller.exe ...
    echo Click "Yes" on the UAC permission prompt.
    echo.
    "%~dp0DeepExcelInstaller.exe"
    goto :done
)

REM Fallback: user-scope PowerShell installer (no admin required)
echo DeepExcelInstaller.exe not found, using PowerShell installer...
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1"

:done
echo.
echo ==========================================
pause
