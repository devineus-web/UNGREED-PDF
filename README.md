# UNGREED-PDF

**Convert any PDF — text-based or scanned/image — into an editable Word document.**

UNGREED-PDF automatically detects whether each page contains extractable text or is a scanned image. Text pages are extracted directly (preserving content), and image pages are processed through OCR (Tesseract) to pull text into the Word document.

---

## Features

- **Text-based PDF** → direct text extraction (fast, high fidelity)
- **Scanned / image PDF** → automatic OCR via Tesseract
- **Mixed PDFs** → handles both types page-by-page
- **Embedded images** → preserved in the Word output
- **Simple GUI** with file picker, progress bar, and one-click conversion
- **CLI mode** for batch/scripting use

## Requirements

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed and on PATH
- [Poppler](https://poppler.freedesktop.org/) (`poppler-utils` / `pdf2image` dependency)

### Install on Ubuntu/Debian
```bash
sudo apt-get install tesseract-ocr poppler-utils
```

### Install on Windows
- Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
- Poppler: https://github.com/oschwartz10612/poppler-windows/releases

### Install on macOS
```bash
brew install tesseract poppler
```

## Setup

```bash
git clone https://github.com/DEVINEUS/UNGREED-PDF.git
cd UNGREED-PDF
pip install -r requirements.txt
```

## Usage

### GUI (Desktop App)
```bash
python app.py
```
A window opens — browse for your PDF, pick an output location, and click **CONVERT TO WORD**.

### CLI
```bash
python cli.py input.pdf                    # output: input.docx
python cli.py input.pdf output.docx        # explicit output path
python cli.py input.pdf --lang eng+fra     # multi-language OCR
```

## Building a Standalone .exe (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name UNGREED-PDF app.py
```
The executable will be in `dist/UNGREED-PDF.exe`.

## How It Works

1. Opens the PDF with PyMuPDF
2. For each page:
   - If extractable text is found → adds text directly to the Word doc
   - If the page is scanned/image → renders at 300 DPI, runs Tesseract OCR, adds extracted text
   - Embedded images are preserved in the output
3. Saves the result as `.docx`

## License

MIT
