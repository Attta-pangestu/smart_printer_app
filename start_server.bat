@echo off
REM Printer Sharing Server Startup Script
REM This script starts the printer sharing server

echo ========================================
echo    Printer Sharing Server Startup
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher
    pause
    exit /b 1
)

REM Check if we're in the correct directory
if not exist "server\main.py" (
    echo ERROR: server\main.py not found
    echo Please run this script from the project root directory
    pause
    exit /b 1
)

REM Install dependencies if requirements.txt exists
if exist "requirements.txt" (
    echo Installing/updating dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo WARNING: Failed to install some dependencies
        echo The server may not work properly
        echo.
    )
)

REM Create logs directory if it doesn't exist
if not exist "logs" mkdir logs

REM Start the server
echo Starting printer sharing server...
echo.
echo Server will be available at:
echo   - Web Interface: http://localhost:8000
echo   - API: http://localhost:8000/api
echo.
echo Press Ctrl+C to stop the server
echo ========================================
echo.

REM Run the server
python -m server.main

REM If we get here, the server has stopped
echo.
echo Server has stopped.
pause