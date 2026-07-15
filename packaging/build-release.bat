@echo off
REM Build the Windows release zip. Run from the repo root on Windows:
REM   packaging\build-release.bat 1.0.0rc3
setlocal enabledelayedexpansion
if "%~1"=="" ( echo usage: packaging\build-release.bat ^<version^> & exit /b 1 )
set VERSION=%~1
cd /d "%~dp0.."

if not exist .venv-build ( py -3 -m venv .venv-build || exit /b 1 )
call .venv-build\Scripts\activate.bat || exit /b 1

REM Old pip cannot install a pyproject-only project; upgrade first, then do a
REM regular (non-editable) install — editable is a dev convenience, not needed
REM to build.
python -m pip install -q --upgrade pip setuptools wheel || exit /b 1
python -m pip install -q . pyinstaller || exit /b 1

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
mkdir dist
python -m networth.generate -o dist\Family_Portfolio_Tracker.xlsx || exit /b 1
cd packaging
pyinstaller --distpath ..\dist --workpath ..\build -y networth-update.spec || ( cd .. & exit /b 1 )
cd ..
if not exist "dist\Update Portfolio.exe" ( echo ERROR: exe was not produced & exit /b 1 )

set STAGE=dist\NetWorth-%VERSION%-windows
mkdir "%STAGE%"
copy dist\Family_Portfolio_Tracker.xlsx "%STAGE%\" >nul
copy "dist\Update Portfolio.exe" "%STAGE%\" >nul
copy packaging\README-enduser.txt "%STAGE%\README.txt" >nul
powershell -NoProfile -Command "Compress-Archive -Path '%STAGE%' -DestinationPath 'dist\NetWorth-%VERSION%-windows.zip' -Force" || exit /b 1
echo Built dist\NetWorth-%VERSION%-windows.zip
