@echo off
REM Double-click to refresh mutual-fund NAVs from AMFI into the tracker.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0UpdateNAV.ps1"
