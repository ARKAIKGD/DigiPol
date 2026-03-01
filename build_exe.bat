@echo off
setlocal

cd /d "%~dp0"
set PYTHONPATH=

if not exist .venv (
	echo Creating virtual environment...
	py -3.12 -m venv .venv
)

set "PYTHON_EXE=%cd%\.venv\Scripts\python.exe"
set "APP_PATH=%cd%\app.py"
set "ICON_PATH=%cd%\assets\studentsnip.ico"
set "BUILD_ROOT=%LOCALAPPDATA%\StudentSnipBuild"
set "DIST_DIR=%BUILD_ROOT%\dist"
set "WORK_DIR=%BUILD_ROOT%\build"
set "SPEC_DIR=%BUILD_ROOT%\spec"
set "PROJECT_DIST=%cd%\dist"

if not exist "%BUILD_ROOT%" mkdir "%BUILD_ROOT%"
if not exist "%SPEC_DIR%" mkdir "%SPEC_DIR%"
if not exist "%PROJECT_DIST%" mkdir "%PROJECT_DIST%"

echo Installing dependencies...
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install -r requirements.txt

echo Generating app icon...
"%PYTHON_EXE%" generate_icon.py

echo Building EXE...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --onefile --windowed --icon "%ICON_PATH%" --name StudentSnip --distpath "%DIST_DIR%" --workpath "%WORK_DIR%" --specpath "%SPEC_DIR%" "%APP_PATH%"

if exist "%DIST_DIR%\StudentSnip.exe" (
	copy /Y "%DIST_DIR%\StudentSnip.exe" "%PROJECT_DIST%\StudentSnip.exe" >nul
)

echo.
echo Done! Your EXE files are here:
echo %DIST_DIR%\StudentSnip.exe
echo %PROJECT_DIST%\StudentSnip.exe
