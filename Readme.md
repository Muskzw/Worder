# Worder: File to Word Converter

A modern Flask web app to convert PDF, image, and text files to Word (.docx) documents.  
Supports OCR for scanned PDFs/images and experimental table extraction from PDFs.  
Includes a glassy, responsive UI with dark mode and a progress bar.

---

## Features

- **Convert PDF, PNG, JPG, JPEG, and TXT to Word (.docx)**
- **OCR support** for scanned PDFs and images (Tesseract)
- **Experimental table extraction** from PDFs (Camelot)
- **Glassy, modern UI** with dark mode toggle
- **Progress bar** and conversion status
- **Keeps original file name** for downloads

---

## Requirements

- Python 3.8+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (for OCR)
- [Ghostscript](https://www.ghostscript.com/) (for Camelot)
- [Poppler](https://github.com/oschwartz10612/poppler-windows/releases/) (for pdf2image)

### Python packages

```bash
pip install flask pdf2docx pdf2image pytesseract pillow python-docx werkzeug camelot-py[cv]
```

---

## Setup

1. **Install Tesseract**  
   - [Download here](https://github.com/tesseract-ocr/tesseract)
   - Add Tesseract to your system PATH.

2. **Install Ghostscript**  
   - [Download here](https://www.ghostscript.com/download/gsdnld.html)
   - Add Ghostscript's `bin` folder to your PATH.

3. **Install Poppler**  
   - [Download here](https://github.com/oschwartz10612/poppler-windows/releases/)
   - Add Poppler's `bin` folder to your PATH.

4. **Install Python dependencies**  
   See above.

---

## Usage

```bash
python app.py
```

- Open [http://127.0.0.1:5000/](http://127.0.0.1:5000/) in your browser.
- Upload your file, select OCR language (for images), and optionally extract tables from PDFs.
- Wait for conversion and download your Word file.

---

## Notes

- **OCR is slower** and may not be 100% accurate, especially for complex layouts.
- **Table extraction** is experimental and may not work for all PDFs.
- Uploaded and converted files are deleted after download for privacy.

---

## Screenshots

![Worder UI Screenshot](screenshot.png)

---

## License

MIT License

---

## Credits

- [Flask](https://flask.palletsprojects.com/)
- [pdf2docx](https://github.com/dothinking/pdf2docx)
- [pytesseract](https://github.com/madmaze/pytesseract)
- [Camelot](https://camelot-py.readthedocs.io/)
-