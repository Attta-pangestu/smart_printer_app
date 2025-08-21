@echo off
REM Printer Sharing Client Startup Script
REM This script starts the printer sharing client GUI

echo ========================================
echo    Printer Sharing Client Startup
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
if not exist "client\gui.py" (
    echo ERROR: client\gui.py not found
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
        echo The client may not work properly
        echo.
    )
)

REM Start the client GUI
echo Starting printer sharing client...
echo.
echo The client GUI will open in a new window
echo Use the GUI to:
echo   - Discover printer servers on your network
echo   - Connect to a specific server by IP
echo   - Print documents with custom settings
echo   - Monitor print jobs
echo.
echo ========================================
echo.

REM Run the client GUI
python -m client.gui

REM If we get here, the client has closed
echo.
echo Client has closed.
pause