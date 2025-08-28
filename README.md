# Pdf-convertor
# PDF to Excel Converter

A simple **Python Tkinter GUI application** that converts PDF files (digital or scanned) into Excel spreadsheets. This tool can handle tables and plain text from PDFs, and provides OCR fallback for scanned documents.

---

## Features

- **Upload PDF**: Select a PDF file from your computer.
- **Convert to Excel**: Extracts tables or text and saves them in an Excel file.
- **OCR Support**: Automatically uses OCR for scanned PDFs without embedded text.
- **Auto Column Widths**: Adjusts column widths in Excel for better readability.
- **User-Friendly GUI**: Simple interface built with Tkinter.

---

## Requirements

- Python 3.8+
- Libraries:
  - `tkinter` (built-in)
  - `pdfplumber`
  - `pytesseract`
  - `pdf2image`
  - `openpyxl`
  - `Pillow (PIL)`

---

## Installation

1. Clone this repository:

```bash
git clone https://dev.azure.com/ark-education/pdf-converter/_git/pdf-converter
cd pdf-converter


2. Install dependencies:

pip install pdfplumber pytesseract pdf2image openpyxl Pillow


3. Install Tesseract OCR:

Windows: Download from Tesseract OCR

Usage

Run the application:

python main.py


In the GUI:

Click "Upload PDF" to select a PDF file.

Click "Convert to Excel" to generate an Excel file.

Click "Exit" to close the app.
