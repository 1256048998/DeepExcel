@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-internal-certificate.ps1"
if errorlevel 1 (
  echo.
  echo Certificate installation failed. Send a screenshot of this window to the DeepExcel developer.
  pause
  exit /b 1
)
echo.
echo Certificate installation completed. You may now run DeepExcel.Setup.INTERNAL.exe.
pause
