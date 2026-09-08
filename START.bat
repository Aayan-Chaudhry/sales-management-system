@echo off
title Sales Tracker
echo.
echo  Starting Sales Tracker...
echo  Opening in your browser at http://localhost:5000
echo.
echo  Keep this window open while using the app.
echo  Close it to stop.
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found.
    echo  Download it from https://python.org and make sure to tick
    echo  "Add Python to PATH" during install.
    pause
    exit
)

:: Install dependencies silently if missing
python -m pip install flask openpyxl --quiet --disable-pip-version-check

:: Backup database before starting the tracker
python backup_db.py
if errorlevel 1 (
    echo.
    echo  WARNING: Database backup failed. The tracker will still start.
    echo  Check the backup message above when you have time.
    echo.
)

:: Run the app
python app.py

pause
