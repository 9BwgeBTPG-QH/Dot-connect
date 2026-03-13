@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

echo ========================================
echo   Dot-connect Setup
echo ========================================
echo.

set "BASE_DIR=%~dp0"
set "PYTHON_DIR=%BASE_DIR%python"
set "PYTHON_VER=3.12.8"
set "PYTHON_ZIP=python-%PYTHON_VER%-embed-amd64.zip"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VER%/%PYTHON_ZIP%"

if exist "%PYTHON_DIR%\python.exe" (
    echo [OK] Python already installed.
    goto install_deps
)

echo [1/4] Downloading Python %PYTHON_VER% ...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%BASE_DIR%%PYTHON_ZIP%'" 2>nul
if errorlevel 1 (
    echo [ERROR] Download failed. Check internet connection.
    pause
    exit /b 1
)

echo [2/4] Extracting...
if not exist "%PYTHON_DIR%" mkdir "%PYTHON_DIR%"
powershell -Command "Expand-Archive -Path '%BASE_DIR%%PYTHON_ZIP%' -DestinationPath '%PYTHON_DIR%' -Force"
del "%BASE_DIR%%PYTHON_ZIP%"

echo [3/4] Setting up pip...
echo import site>> "%PYTHON_DIR%\python312._pth"
powershell -Command "Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%PYTHON_DIR%\get-pip.py'" 2>nul
"%PYTHON_DIR%\python.exe" "%PYTHON_DIR%\get-pip.py" --no-warn-script-location -q
del "%PYTHON_DIR%\get-pip.py"

:install_deps
echo [4/5] Installing dependencies...
"%PYTHON_DIR%\python.exe" -m pip install -r "%BASE_DIR%requirements.txt" --no-warn-script-location -q
"%PYTHON_DIR%\python.exe" -m pip install pywin32 --no-warn-script-location -q

echo [5/5] Registering pywin32 COM components...
"%PYTHON_DIR%\python.exe" "%PYTHON_DIR%\Scripts\pywin32_postinstall.py" -install 2>nul
if errorlevel 1 (
    "%PYTHON_DIR%\python.exe" -c "import pywin32_postinstall; pywin32_postinstall.install()" 2>nul
)

echo.
echo ========================================
echo   Setup complete!
echo   Double-click start.bat to launch.
echo ========================================
pause
