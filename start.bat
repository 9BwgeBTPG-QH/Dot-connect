@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "PYTHON_DIR=%~dp0python"

if not exist "%PYTHON_DIR%\python.exe" (
    echo Python not found. Please run setup.bat first.
    pause
    exit /b 1
)

echo ========================================
echo   Dot-connect
echo   http://localhost:8000
echo   Log file: server.log
echo   Close this window to stop the server.
echo ========================================
echo.

start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:8000"

"%PYTHON_DIR%\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8000
