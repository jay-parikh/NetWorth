@echo off
REM Double-click to refresh the MF_Master fund list (AMFI) in the tracker.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0UpdateFundMaster.ps1"
