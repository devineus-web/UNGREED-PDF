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
- **No Poppler required** — uses PyMuPDF for native PDF rendering
- **Auto-detects Tesseract** on Windows (no PATH setup needed)

## Requirements

- **Windows:** Just install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) — run the installer, keep default path. That's it.
- **macOS:** `brew install tesseract`
- **Linux:** `sudo apt-get install tesseract-ocr`

## Quick Start (Windows)

1. Download `UNGREED-PDF.exe` from the [Releases page](https://github.com/devineus-web/UNGREED-PDF/releases)
2. Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) (just run the installer)
3. Double-click `UNGREED-PDF.exe`
4. Click **Browse** to select your PDF → Click **CONVERT TO WORD**

## Running from Source

```bash
git clone https://github.com/devineus-web/UNGREED-PDF.git
cd UNGREED-PDF
pip install -r requirements.txt
python app.py
```

## CLI

```bash
python cli.py input.pdf                    # output: input.docx
python cli.py input.pdf output.docx        # explicit output path
python cli.py input.pdf --lang eng+fra     # multi-language OCR
```

## Building the .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name UNGREED-PDF --add-data "converter.py;." app.py
```

## How It Works

1. Opens the PDF with PyMuPDF
2. For each page:
   - If extractable text is found → adds text directly to the Word doc
   - If the page is scanned/image → renders at 300 DPI with PyMuPDF, runs Tesseract OCR, adds extracted text
   - Embedded images are preserved in the output
3. Saves the result as `.docx`

## License

MIT
