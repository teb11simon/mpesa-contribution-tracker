@echo off
REM M-Pesa Contribution Tracker - Windows Launcher

echo ========================================
echo   M-Pesa Contribution Tracker
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

REM Check if Tesseract is installed
tesseract --version >nul 2>&1
if %errorlevel% neq 0 (
    echo WARNING: Tesseract OCR is not installed or not in PATH
    echo OCR functionality may not work properly
    echo.
    echo To install Tesseract:
    echo 1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
    echo 2. Install and add to PATH: C:\Program Files\Tesseract-OCR
    echo.
    pause
)

REM Check if dependencies are installed
python -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

REM Run the application
echo.
echo Starting application...
echo.
python src\main.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Application failed to start
    pause
    exit /b 1
)
