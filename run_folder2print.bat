@echo off
REM Folder2Print - Automatic PDF Printing
REM This batch file runs the folder2print script

cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

REM Run the script
python folder2print.py

REM If the script exits, pause to see any error messages
pause
