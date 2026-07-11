@echo off
REM DeepExcel Installer - double-click to run
REM This wrapper calls register-user.ps1 with execution policy bypassed
REM so the user doesn't need to change PowerShell settings.

REM Change to the directory containing this .bat file
cd /d "%~dp0"

REM Run the PowerShell registration script
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0register-user.ps1"

REM The PowerShell script has its own "Press Enter" prompt, but keep this
REM pause as a fallback in case the PS script exits unexpectedly.
pause
