@echo off
cd /d "%~dp0"
echo.
echo ============================================================
echo   File Merger Pro — Setup and Launch
echo ============================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python is not installed or not in PATH.
    echo  Download from: https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  Installing / updating required packages ...
pip install -r requirements.txt --quiet --upgrade
if errorlevel 1 (
    echo.
    echo  ERROR: pip install failed. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo  Starting File Merger Pro ...
echo  Open your browser at:  http://localhost:8501
echo  Press Ctrl+C in this window to stop the app.
echo.
streamlit run file_merger_app.py --server.maxUploadSize 500
