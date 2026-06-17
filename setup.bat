@echo off
REM Setup Script for M-Pesa Contribution Tracker

echo ========================================
echo   M-Pesa Contribution Tracker Setup
echo ========================================
echo.

REM Check Python installation
echo [1/4] Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)
python --version
echo Python found!
echo.

REM Check Tesseract installation
echo [2/4] Checking Tesseract OCR installation...
tesseract --version >nul 2>&1
if %errorlevel% neq 0 (
    echo WARNING: Tesseract OCR is not installed
    echo.
    echo To install Tesseract OCR:
    echo 1. Visit: https://github.com/UB-Mannheim/tesseract/wiki
    echo 2. Download the latest Windows installer (tesseract-ocr-w64-setup-*.exe)
    echo 3. Run the installer
    echo 4. Add to PATH: C:\Program Files\Tesseract-OCR
    echo.
    echo Press any key to continue without Tesseract (OCR will not work)...
    pause >nul
) else (
    tesseract --version | findstr /C:"tesseract"
    echo Tesseract found!
)
echo.

REM Install Python dependencies
echo [3/4] Installing Python dependencies...
pip install --upgrade pip
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo Dependencies installed successfully!
echo.

REM Create output directory
echo [4/4] Creating output directory...
if not exist "output" mkdir output
echo Output directory created!
echo.

REM Done
echo ========================================
echo   Setup Complete!
echo ========================================
echo.
echo You can now run the application by:
echo   1. Double-clicking run.bat
echo   2. Or running: python src\main.py
echo.
echo For more information, see README.md
echo.
pause
