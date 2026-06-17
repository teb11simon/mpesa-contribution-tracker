# M-Pesa Contribution Tracker

A Windows desktop application that processes M-Pesa statements and contribution notes images to generate Excel contribution reports.

## Features

- **M-Pesa Statement Parsing**: Extract transactions from PDF or CSV M-Pesa statements
- **OCR Processing**: Extract contribution data from handwritten notes images
- **Excel Generation**: Create formatted Excel reports with multiple sheets
- **Image Enhancement**: Improve OCR accuracy with image preprocessing
- **Categorization**: Automatically categorize contributions (Contribution, Benevolence, Missions)

## Requirements

- Python 3.8 or higher
- Windows 10 or later

## Installation

### 1. Install Python

Download and install Python from [python.org](https://www.python.org/downloads/)

Make sure to check "Add Python to PATH" during installation.

### 2. Install Tesseract OCR (Required for local OCR)

Download and install Tesseract OCR for Windows:

1. Go to [UB Mannheim Tesseract Wiki](https://github.com/UB-Mannheim/tesseract/wiki)
2. Download the latest Windows installer (e.g., `tesseract-ocr-w64-setup-5.x.x.exe`)
3. Run the installer and complete the installation
4. Add Tesseract to your PATH (usually: `C:\Program Files\Tesseract-OCR`)

### 3. Install Python Dependencies

Open Command Prompt or PowerShell and run:

```bash
cd "c:\Users\User\Documents\Personal project\mpesa-contribution-tracker"
pip install -r requirements.txt
```

### 4. (Optional) Google Cloud Vision API

For better OCR accuracy, you can use Google Cloud Vision API instead of Tesseract:

1. Create a Google Cloud project
2. Enable the Vision API
3. Create a service account and download the JSON key file
4. Set the environment variable:
   ```bash
   set GOOGLE_APPLICATION_CREDENTIALS=path\to\your\key.json
   ```

## Usage

### Running the Application

```bash
cd "c:\Users\User\Documents\Personal project\mpesa-contribution-tracker"
python src/main.py
```

### Using the App

1. **Upload M-Pesa Statement**: Click "Browse..." to select your M-Pesa statement (PDF or CSV)
2. **Upload Contribution Notes**: Click "Browse..." to select an image of handwritten contribution notes
3. **Set Options**:
   - Choose the report date
   - Optionally enable Google Vision API for better OCR
   - Enable image enhancement for better OCR accuracy
4. **Generate Report**: Click "Generate Excel Report"
5. **View Results**: The app will process the files and generate an Excel report

### Supported Image Formats

- PNG
- JPG/JPEG
- BMP
- TIFF

### Supported M-Pesa Statement Formats

- PDF
- CSV

## Project Structure

```
mpesa-contribution-tracker/
├── src/
│   ├── __init__.py
│   ├── main.py              # Main GUI application
│   ├── mpesa_parser.py      # M-Pesa statement parser
│   ├── ocr_processor.py     # OCR processing module
│   └── excel_generator.py   # Excel report generator
├── output/                  # Generated Excel reports
├── requirements.txt         # Python dependencies
└── README.md               # This file
```

## Troubleshooting

### Tesseract Not Found

If you get an error about Tesseract not being found:

1. Make sure Tesseract is installed
2. Add it to your Windows PATH:
   - Search for "Environment Variables" in Windows
   - Edit "Path" under System Variables
   - Add: `C:\Program Files\Tesseract-OCR`
3. Restart Command Prompt/PowerShell

### PDF Parsing Issues

If PDF parsing fails:

1. Make sure `pdfplumber` is installed: `pip install pdfplumber`
2. Try converting the PDF to CSV first and use the CSV parser

### OCR Accuracy

For better OCR results:

1. Use clear, well-lit photos
2. Enable "Enhance image for better OCR" option
3. Consider using Google Vision API (requires API key)
4. Write names and amounts clearly

## License

MIT License

## Support

For issues or questions, please contact the development team.
