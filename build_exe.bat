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

echo Closing any running StudentSnip instances...
taskkill /IM StudentSnip.exe /F >nul 2>&1

echo Installing dependencies...
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install -r requirements.txt

echo Generating app icon...
"%PYTHON_EXE%" generate_icon.py

echo Building EXE...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --onefile --windowed --icon "%ICON_PATH%" --add-data "%ICON_PATH%;assets" --name StudentSnip --copy-metadata imageio --copy-metadata imageio-ffmpeg --hidden-import imageio --hidden-import imageio.v2 --hidden-import imageio_ffmpeg --hidden-import numpy --distpath "%DIST_DIR%" --workpath "%WORK_DIR%" --specpath "%SPEC_DIR%" "%APP_PATH%"

if errorlevel 1 (
    echo Build failed.
    exit /b 1
)

if exist "%PROJECT_DIST%\StudentSnip.exe" del /F /Q "%PROJECT_DIST%\StudentSnip.exe" >nul 2>&1

if exist "%DIST_DIR%\StudentSnip.exe" (
	copy /Y "%DIST_DIR%\StudentSnip.exe" "%PROJECT_DIST%\StudentSnip.exe" >nul
) else (
	echo Build output missing: %DIST_DIR%\StudentSnip.exe
	exit /b 1
)

if not exist "%PROJECT_DIST%\StudentSnip.exe" (
	echo Failed to refresh project dist exe.
	exit /b 1
)

echo.
echo Done! Your EXE files are here:
echo %DIST_DIR%\StudentSnip.exe
echo %PROJECT_DIST%\StudentSnip.exe
