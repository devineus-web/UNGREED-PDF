"""
UNGREED-PDF — Desktop GUI Application
A simple, clean interface for converting PDFs to editable Word documents.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from converter import convert_pdf_to_docx


# ──────────────────────────────────────────────
# Theme / Colors
# ──────────────────────────────────────────────
BG_DARK = "#1a1a2e"
BG_CARD = "#16213e"
ACCENT = "#0f3460"
HIGHLIGHT = "#e94560"
TEXT_PRIMARY = "#ffffff"
TEXT_SECONDARY = "#a0a0b0"
SUCCESS_GREEN = "#00c853"


class UngreedPDFApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("UNGREED-PDF")
        self.root.geometry("700x520")
        self.root.minsize(600, 480)
        self.root.configure(bg=BG_DARK)
        self.root.resizable(True, True)

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready — select a PDF to convert")
        self.is_converting = False

        self._build_ui()

    # ── UI Construction ──────────────────────
    def _build_ui(self):
        # Title bar
        title_frame = tk.Frame(self.root, bg=HIGHLIGHT, height=60)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)

        tk.Label(
            title_frame,
            text="UNGREED-PDF",
            font=("Segoe UI", 22, "bold"),
            bg=HIGHLIGHT,
            fg=TEXT_PRIMARY,
        ).pack(side="left", padx=20, pady=10)

        tk.Label(
            title_frame,
            text="PDF → Word Converter",
            font=("Segoe UI", 11),
            bg=HIGHLIGHT,
            fg="#ffcdd2",
        ).pack(side="left", pady=10)

        # Main content
        content = tk.Frame(self.root, bg=BG_DARK, padx=30, pady=20)
        content.pack(fill="both", expand=True)

        # ── Input PDF ─────────────────────
        self._section_label(content, "INPUT PDF")

        input_frame = tk.Frame(content, bg=BG_CARD, relief="flat", bd=0)
        input_frame.pack(fill="x", pady=(0, 15))

        input_inner = tk.Frame(input_frame, bg=BG_CARD, padx=15, pady=12)
        input_inner.pack(fill="x")

        self.input_entry = tk.Entry(
            input_inner,
            textvariable=self.pdf_path,
            font=("Consolas", 10),
            bg="#0d1b2a",
            fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief="flat",
            bd=0,
            state="readonly",
            readonlybackground="#0d1b2a",
            cursor="arrow",
        )
        self.input_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.input_entry.bind("<Button-1>", lambda e: self._browse_pdf())

        browse_btn = tk.Button(
            input_inner,
            text="Browse",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT,
            fg=TEXT_PRIMARY,
            activebackground=HIGHLIGHT,
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=4,
            command=self._browse_pdf,
        )
        browse_btn.pack(side="right", padx=(10, 0))

        # ── Output Location ───────────────
        self._section_label(content, "OUTPUT LOCATION (optional)")

        output_frame = tk.Frame(content, bg=BG_CARD, relief="flat", bd=0)
        output_frame.pack(fill="x", pady=(0, 15))

        output_inner = tk.Frame(output_frame, bg=BG_CARD, padx=15, pady=12)
        output_inner.pack(fill="x")

        self.output_entry = tk.Entry(
            output_inner,
            textvariable=self.output_path,
            font=("Consolas", 10),
            bg="#0d1b2a",
            fg=TEXT_PRIMARY,
            insertbackground=TEXT_PRIMARY,
            relief="flat",
            bd=0,
            state="readonly",
            readonlybackground="#0d1b2a",
            cursor="arrow",
        )
        self.output_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.output_entry.bind("<Button-1>", lambda e: self._browse_output())

        out_btn = tk.Button(
            output_inner,
            text="Browse",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT,
            fg=TEXT_PRIMARY,
            activebackground=HIGHLIGHT,
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=4,
            command=self._browse_output,
        )
        out_btn.pack(side="right", padx=(10, 0))

        # ── Progress ──────────────────────
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=BG_CARD,
            background=HIGHLIGHT,
            thickness=18,
        )

        self.progress = ttk.Progressbar(
            content,
            style="Custom.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
        )
        self.progress.pack(fill="x", pady=(5, 8))

        self.progress_label = tk.Label(
            content,
            text="",
            font=("Segoe UI", 9),
            bg=BG_DARK,
            fg=TEXT_SECONDARY,
        )
        self.progress_label.pack(anchor="w")

        # ── Convert Button ────────────────
        self.convert_btn = tk.Button(
            content,
            text="⚡  CONVERT TO WORD",
            font=("Segoe UI", 14, "bold"),
            bg=HIGHLIGHT,
            fg=TEXT_PRIMARY,
            activebackground="#c62828",
            activeforeground=TEXT_PRIMARY,
            relief="flat",
            cursor="hand2",
            pady=10,
            command=self._start_conversion,
        )
        self.convert_btn.pack(fill="x", pady=(15, 10))

        # ── Status Bar ────────────────────
        status_bar = tk.Frame(self.root, bg=ACCENT, height=32)
        status_bar.pack(fill="x", side="bottom")
        status_bar.pack_propagate(False)

        tk.Label(
            status_bar,
            textvariable=self.status_text,
            font=("Segoe UI", 9),
            bg=ACCENT,
            fg=TEXT_SECONDARY,
            anchor="w",
        ).pack(fill="x", padx=15, pady=6)

    def _section_label(self, parent, text):
        tk.Label(
            parent,
            text=text,
            font=("Segoe UI", 9, "bold"),
            bg=BG_DARK,
            fg=TEXT_SECONDARY,
        ).pack(anchor="w", pady=(8, 4))

    # ── Actions ──────────────────────────────
    def _browse_pdf(self):
        initial = os.path.dirname(self.pdf_path.get()) if self.pdf_path.get() else os.path.expanduser("~")
        path = filedialog.askopenfilename(
            title="Select PDF",
            initialdir=initial,
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            # Always auto-fill output: same folder, same name but .docx
            base = os.path.splitext(path)[0]
            self.output_path.set(base + ".docx")

    def _browse_output(self):
        # Start in the same folder as the selected PDF, with a suggested name
        current_pdf = self.pdf_path.get()
        if current_pdf:
            initial_dir = os.path.dirname(current_pdf)
            initial_file = os.path.splitext(os.path.basename(current_pdf))[0] + ".docx"
        else:
            initial_dir = os.path.expanduser("~")
            initial_file = "converted.docx"

        path = filedialog.asksaveasfilename(
            title="Save Word Document As",
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _start_conversion(self):
        pdf = self.pdf_path.get().strip()
        if not pdf:
            messagebox.showwarning("No PDF Selected", "Please select a PDF file first.")
            return
        if not os.path.isfile(pdf):
            messagebox.showerror("File Not Found", f"Cannot find:\n{pdf}")
            return
        if self.is_converting:
            return

        self.is_converting = True
        self.convert_btn.configure(state="disabled", text="Converting…")
        self.progress["value"] = 0
        self.progress_label.configure(text="")
        self.status_text.set("Converting — please wait…")

        output = self.output_path.get().strip() or None
        thread = threading.Thread(
            target=self._run_conversion, args=(pdf, output), daemon=True
        )
        thread.start()

    def _run_conversion(self, pdf_path, output_path):
        try:
            result = convert_pdf_to_docx(
                pdf_path,
                output_path,
                progress_callback=self._on_progress,
            )
            self.root.after(0, self._on_success, result)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_progress(self, current, total):
        pct = int((current / total) * 100)
        self.root.after(0, self._update_progress, pct, current, total)

    def _update_progress(self, pct, current, total):
        self.progress["value"] = pct
        self.progress_label.configure(text=f"Page {current} of {total}  ({pct}%)")

    def _on_success(self, output_path):
        self.is_converting = False
        self.convert_btn.configure(state="normal", text="⚡  CONVERT TO WORD")
        self.progress["value"] = 100
        self.status_text.set(f"Done — saved to {output_path}")
        self.progress_label.configure(text="Conversion complete!")
        messagebox.showinfo(
            "Conversion Complete",
            f"Your Word document has been saved to:\n\n{output_path}",
        )

    def _on_error(self, error_msg):
        self.is_converting = False
        self.convert_btn.configure(state="normal", text="⚡  CONVERT TO WORD")
        self.progress["value"] = 0
        self.status_text.set("Error during conversion")
        self.progress_label.configure(text="")
        messagebox.showerror("Conversion Failed", f"An error occurred:\n\n{error_msg}")


def main():
    root = tk.Tk()
    app = UngreedPDFApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
