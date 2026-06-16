@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Create-AccDropdownAttributes.ps1"
echo.
pause
