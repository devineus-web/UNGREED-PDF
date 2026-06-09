#!/usr/bin/env python3
"""
UNGREED-PDF — Command-Line Interface
Usage:  python cli.py input.pdf [output.docx] [--lang eng]
"""

import argparse
import sys
import os
from converter import convert_pdf_to_docx


def main():
    parser = argparse.ArgumentParser(
        prog="ungreed-pdf",
        description="UNGREED-PDF: Convert any PDF (text or scanned) to editable Word .docx",
    )
    parser.add_argument("pdf", help="Path to the input PDF file")
    parser.add_argument("output", nargs="?", default=None, help="Path for the output .docx (optional)")
    parser.add_argument("--lang", default="eng", help="Tesseract OCR language (default: eng)")
    args = parser.parse_args()

    if not os.path.isfile(args.pdf):
        print(f"Error: file not found — {args.pdf}", file=sys.stderr)
        sys.exit(1)

    def progress(current, total):
        bar_len = 40
        filled = int(bar_len * current / total)
        bar = "█" * filled + "░" * (bar_len - filled)
        pct = int(100 * current / total)
        print(f"\r  [{bar}] {pct}%  page {current}/{total}", end="", flush=True)

    print(f"UNGREED-PDF — Converting: {args.pdf}")
    result = convert_pdf_to_docx(args.pdf, args.output, progress_callback=progress, ocr_lang=args.lang)
    print(f"\n  ✓ Saved to: {result}")


if __name__ == "__main__":
    main()
