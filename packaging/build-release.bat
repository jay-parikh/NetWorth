@echo off
REM Build the Windows release zip. Run from the repo root:
REM   packaging\build-release.bat 0.4.0
setlocal
if "%~1"=="" ( echo usage: packaging\build-release.bat ^<version^> & exit /b 1 )
set VERSION=%~1
cd /d "%~dp0.."

if not exist .venv-build ( py -3 -m venv .venv-build )
call .venv-build\Scripts\activate.bat
pip -q install -e . pyinstaller

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
mkdir dist
python -m networth.generate -o dist\Family_Portfolio_Tracker.xlsx
cd packaging
pyinstaller --distpath ..\dist --workpath ..\build -y networth-update.spec
cd ..

set STAGE=dist\NetWorth-%VERSION%-windows
mkdir "%STAGE%"
copy dist\Family_Portfolio_Tracker.xlsx "%STAGE%\" >nul
copy "dist\Update Portfolio.exe" "%STAGE%\" >nul
copy packaging\README-enduser.txt "%STAGE%\README.txt" >nul
powershell -Command "Compress-Archive -Path '%STAGE%' -DestinationPath 'dist\NetWorth-%VERSION%-windows.zip' -Force"
echo Built dist\NetWorth-%VERSION%-windows.zip
