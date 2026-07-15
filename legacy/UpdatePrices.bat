@echo off
cd /d "%~dp0"
echo Running price update... (a PowerShell window will guide you)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0UpdatePrices.ps1"
echo.
echo If the PowerShell window closed too fast, open UpdatePrices_log.txt in this folder.
pause
