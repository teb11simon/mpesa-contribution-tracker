# Quick Start Guide

## First Time Setup (5 minutes)

### Step 1: Install Python (if not already installed)
1. Download from https://www.python.org/downloads/
2. Run the installer
3. **Important**: Check "Add Python to PATH"
4. Click "Install Now"

### Step 2: Install Tesseract OCR (for image processing)
1. Download from https://github.com/UB-Mannheim/tesseract/wiki
2. Look for "tesseract-ocr-w64-setup-5.x.x.exe"
3. Run the installer
4. Add to PATH: `C:\Program Files\Tesseract-OCR`

### Step 3: Run Setup
1. Double-click `setup.bat`
2. Wait for installation to complete
3. Press any key to finish

## Running the App

### Option 1: Double-click `run.bat`
This will launch the application directly.

### Option 2: From Command Prompt
```bash
cd "c:\Users\User\Documents\Personal project\mpesa-contribution-tracker"
python src\main.py
```

## Using the App

1. **Upload M-Pesa Statement**
   - Click "Browse..." next to "M-Pesa Statement"
   - Select your PDF or CSV file

2. **Upload Contribution Notes**
   - Click "Browse..." next to "Contribution Notes Image"
   - Select an image (PNG, JPG, etc.)

3. **Configure Options**
   - Set the report date
   - Enable "Enhance image" for better OCR (recommended)
   - Enable "Google Vision API" if you have it set up

4. **Generate Report**
   - Click "Generate Excel Report"
   - Wait for processing
   - Click "Open File" to view the report

## Tips for Best Results

### For M-Pesa Statements
- Use the official M-Pesa statement PDF
- Ensure the statement is not password protected
- CSV format works best if available

### For Contribution Notes Images
- Take photos in good lighting
- Write names and amounts clearly
- Use dark ink on light paper
- Keep the camera steady
- Enable "Enhance image" option

### Expected Format for Handwritten Notes
The OCR looks for patterns like:
- `John Doe - 5000`
- `Jane Smith: 2500 Contribution`
- `5000 - John Doe (Missions)`
- `Bob Johnson 1000 Benevolence`

## Troubleshooting

### "Tesseract not found" error
1. Make sure Tesseract is installed
2. Add to PATH: `C:\Program Files\Tesseract-OCR`
3. Restart Command Prompt

### "Python not found" error
1. Reinstall Python
2. Check "Add Python to PATH" during installation
3. Restart Command Prompt

### OCR not recognizing text
1. Enable "Enhance image" option
2. Take a clearer photo
3. Consider using Google Vision API

### PDF parsing fails
1. Try converting PDF to CSV first
2. Make sure PDF is not password protected
3. Check that `pdfplumber` is installed

## Getting Help

For more detailed information, see [README.md](README.md)
