@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0AccAttributeTool.ps1" -Mode populate-file-attributes-from-folders -DryRun
pause
