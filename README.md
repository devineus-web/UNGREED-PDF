# UNGREED-PDF

**Convert any PDF — text-based or scanned/image — into an editable Word document.**

Works out of the box on Windows 10/11. No extra software to install.

---

## Quick Start (Windows)

1. Download `UNGREED-PDF.exe` from the [Releases page](https://github.com/devineus-web/UNGREED-PDF/releases)
2. Double-click it
3. Click **Browse** to select your PDF → Click **CONVERT TO WORD**

That's it. Tesseract OCR is bundled inside the `.exe`.

## Features

- **Text-based PDF** → direct text extraction (fast, high fidelity)
- **Scanned / image PDF** → automatic OCR (Tesseract bundled)
- **Mixed PDFs** → handles both types page-by-page
- **Embedded images** → preserved in the Word output
- **Simple GUI** with file picker, progress bar, and one-click conversion
- **CLI mode** for batch/scripting use
- **Zero dependencies** — everything is bundled in the `.exe`

## Running from Source

If you prefer running from source:

```bash
git clone https://github.com/devineus-web/UNGREED-PDF.git
cd UNGREED-PDF
pip install -r requirements.txt
python app.py
```

When running from source, you'll need Tesseract OCR installed:
- **Windows:** https://github.com/UB-Mannheim/tesseract/wiki
- **macOS:** `brew install tesseract`
- **Linux:** `sudo apt-get install tesseract-ocr`

## CLI

```bash
python cli.py input.pdf                    # output: input.docx
python cli.py input.pdf output.docx        # explicit output path
python cli.py input.pdf --lang eng+fra     # multi-language OCR
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
