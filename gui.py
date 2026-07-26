"""
Desktop GUI - Arabic PDF to Excel Extractor
Created By : Jawaz Ali
Contact    : +966-0539618563
"""

import os
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    import ttkbootstrap as ttk
    HAS_BOOTSTRAP = True
except Exception:
    import tkinter.ttk as ttk
    HAS_BOOTSTRAP = False

from utils.pdf_analyzer import PDFAnalyzer
from core.text_extractor import TextExtractor
from core.ocr_extractor import OCRExtractor
from core.table_extractor import TableExtractor
from core.invoice_extractor import InvoiceExtractor
from core.batch_processor import BatchProcessor


# ══════════════════════════════════════════════════════════
#  App Info
# ══════════════════════════════════════════════════════════
APP_NAME    = "Arabic PDF to Excel Extractor"
APP_VERSION = "1.0.0"
CREATOR     = "Jawaz Ali"
CONTACT     = "+966-0539618563"
YEAR        = "2025"


# ══════════════════════════════════════════════════════════
#  Window Factory
# ══════════════════════════════════════════════════════════
def _make_root():
    if HAS_BOOTSTRAP:
        try:
            root = ttk.Window(
                title=f"{APP_NAME} — by {CREATOR}",
                themename="cosmo",
                size=(980, 800),
                resizable=(True, True),
            )
            root.minsize(820, 640)
            return root
        except Exception:
            pass
    root = tk.Tk()
    root.title(f"{APP_NAME} — by {CREATOR}")
    root.geometry("980x800")
    root.minsize(820, 640)
    return root


# ══════════════════════════════════════════════════════════
#  Main Application
# ══════════════════════════════════════════════════════════
class PDFExtractorApp:
    """Main GUI — Arabic PDF to Excel Extractor by Jawaz Ali"""

    def __init__(self):
        self.root = _make_root()
        self.pdf_files = []
        self.is_processing = tk.BooleanVar(value=False)
        self.single_file_var  = tk.StringVar()
        self.output_dir_var   = tk.StringVar(
            value=str(Path.home() / "Desktop" / "PDF_Extracted")
        )
        self.method_var      = tk.StringVar(value="auto")
        self.ocr_engine_var  = tk.StringVar(value="tesseract")
        self.table_mode_var  = tk.StringVar(value="separate")
        self._build_ui()

    # ──────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=10, pady=5)
        self._tab_single(nb)
        self._tab_batch(nb)
        self._tab_settings(nb)
        self._tab_about(nb)
        self._build_statusbar()

    # ══════════════════════════════════════════════════════
    #  Header
    # ══════════════════════════════════════════════════════
    def _build_header(self):
        # ── App title ──
        title_frame = ttk.Frame(self.root, padding=(10, 8))
        title_frame.pack(fill="x")

        ttk.Label(
            title_frame,
            text=f"  {APP_NAME}",
            font=("Segoe UI", 17, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            title_frame,
            text="  Text PDF  |  Scanned OCR  |  Tables  |  Invoices  |  Batch  —  Arabic preserved unchanged",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(2, 0))

        # ── Creator info bar ──
        bar = tk.Frame(self.root, bg="#1a3a6b", pady=6)
        bar.pack(fill="x")

        left = tk.Frame(bar, bg="#1a3a6b")
        left.pack(side="left", padx=15)

        tk.Label(
            left,
            text=f"  Created By:  {CREATOR}",
            font=("Segoe UI", 10, "bold"),
            fg="#ffffff",
            bg="#1a3a6b",
        ).pack(side="left")

        tk.Label(
            left,
            text="   |   ",
            font=("Segoe UI", 10),
            fg="#aac4ff",
            bg="#1a3a6b",
        ).pack(side="left")

        tk.Label(
            left,
            text=f"Contact:  {CONTACT}",
            font=("Segoe UI", 10, "bold"),
            fg="#ffd966",
            bg="#1a3a6b",
        ).pack(side="left")

        tk.Label(
            left,
            text="   |   ",
            font=("Segoe UI", 10),
            fg="#aac4ff",
            bg="#1a3a6b",
        ).pack(side="left")

        tk.Label(
            left,
            text=f"Version {APP_VERSION}  —  {YEAR}",
            font=("Segoe UI", 9),
            fg="#aac4ff",
            bg="#1a3a6b",
        ).pack(side="left")

        ttk.Separator(self.root, orient="horizontal").pack(fill="x")

    # ══════════════════════════════════════════════════════
    #  Status Bar
    # ══════════════════════════════════════════════════════
    def _build_statusbar(self):
        ttk.Separator(self.root, orient="horizontal").pack(fill="x", side="bottom")

        bar = tk.Frame(self.root, bg="#f0f0f0", pady=3)
        bar.pack(fill="x", side="bottom")

        self._global_status = tk.StringVar(value="Ready — select a PDF file to begin")

        tk.Label(
            bar,
            textvariable=self._global_status,
            font=("Segoe UI", 9),
            anchor="w",
            bg="#f0f0f0",
            fg="#333333",
        ).pack(side="left", padx=10)

        tk.Label(
            bar,
            text=f"Created by {CREATOR}  |  {CONTACT}",
            font=("Segoe UI", 8),
            anchor="e",
            bg="#f0f0f0",
            fg="#888888",
        ).pack(side="right", padx=10)

    # ══════════════════════════════════════════════════════
    #  TAB 1 — Single File
    # ══════════════════════════════════════════════════════
    def _tab_single(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Single File   ")

        # File row
        fg = ttk.LabelFrame(tab, text=" PDF File ", padding=8)
        fg.pack(fill="x", pady=4)

        ttk.Entry(
            fg,
            textvariable=self.single_file_var,
            font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))

        ttk.Button(fg, text="Browse", command=self._browse_single, width=10).pack(side="left", padx=2)
        ttk.Button(fg, text="Analyze", command=self._analyze_pdf, width=10).pack(side="left", padx=2)

        # Analysis box
        af = ttk.LabelFrame(tab, text=" PDF Analysis ", padding=6)
        af.pack(fill="x", pady=4)

        self._analysis_box = ScrolledText(
            af,
            height=8,
            font=("Consolas", 10),
            bg="#f8f9fa",
            relief="flat",
            state="disabled",
            wrap="word",
        )
        self._analysis_box.pack(fill="x")

        # Method
        mf = ttk.LabelFrame(tab, text=" Extraction Method ", padding=8)
        mf.pack(fill="x", pady=4)

        methods = [
            ("auto",    "Auto Detect"),
            ("text",    "Text PDF"),
            ("ocr",     "OCR Scanned"),
            ("table",   "Tables Only"),
            ("invoice", "Invoice/Form"),
        ]
        for val, lbl in methods:
            ttk.Radiobutton(mf, text=lbl, variable=self.method_var, value=val).pack(
                side="left", padx=10)

        # Output dir
        of = ttk.LabelFrame(tab, text=" Output Directory ", padding=8)
        of.pack(fill="x", pady=4)

        ttk.Entry(of, textvariable=self.output_dir_var, font=("Segoe UI", 10)).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(of, text="Browse", command=self._browse_output, width=10).pack(side="left")

        # Extract button
        bf = ttk.Frame(tab)
        bf.pack(pady=8)

        self._extract_btn = ttk.Button(
            bf,
            text="Extract to Excel",
            command=self._do_extract_single,
            width=30,
        )
        self._extract_btn.pack()

        # Progress
        self._single_bar = ttk.Progressbar(tab, orient="horizontal", mode="determinate")
        self._single_bar.pack(fill="x", pady=3)

        self._single_status = tk.StringVar(value="Idle")
        ttk.Label(tab, textvariable=self._single_status, font=("Segoe UI", 9)).pack()

    # ══════════════════════════════════════════════════════
    #  TAB 2 — Batch
    # ══════════════════════════════════════════════════════
    def _tab_batch(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Batch Process   ")

        br = ttk.Frame(tab)
        br.pack(fill="x", pady=(0, 6))

        ttk.Button(br, text="Add Files",   command=self._batch_add_files,  width=14).pack(side="left", padx=3)
        ttk.Button(br, text="Add Folder",  command=self._batch_add_folder, width=14).pack(side="left", padx=3)
        ttk.Button(br, text="Clear All",   command=self._batch_clear,      width=14).pack(side="left", padx=3)

        lf = ttk.LabelFrame(tab, text=" Files ", padding=5)
        lf.pack(fill="both", expand=True, pady=4)

        cols = ("name", "size", "status")
        self._tree = ttk.Treeview(lf, columns=cols, show="headings", height=10)
        self._tree.heading("name",   text="File Name")
        self._tree.heading("size",   text="Size")
        self._tree.heading("status", text="Status")
        self._tree.column("name",   width=430, anchor="w")
        self._tree.column("size",   width=90,  anchor="center")
        self._tree.column("status", width=160, anchor="center")

        vs = ttk.Scrollbar(lf, orient="vertical",   command=self._tree.yview)
        hs = ttk.Scrollbar(lf, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        lf.grid_rowconfigure(0, weight=1)
        lf.grid_columnconfigure(0, weight=1)

        self._batch_btn = ttk.Button(
            tab,
            text="Process All Files",
            command=self._do_batch,
            width=30,
        )
        self._batch_btn.pack(pady=8)

        self._batch_bar = ttk.Progressbar(tab, orient="horizontal", mode="determinate")
        self._batch_bar.pack(fill="x", pady=3)

        self._batch_status = tk.StringVar(value="No files added")
        ttk.Label(tab, textvariable=self._batch_status, font=("Segoe UI", 9)).pack()

    # ══════════════════════════════════════════════════════
    #  TAB 3 — Settings
    # ══════════════════════════════════════════════════════
    def _tab_settings(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Settings   ")

        # OCR Engine
        of = ttk.LabelFrame(tab, text=" OCR Engine ", padding=10)
        of.pack(fill="x", pady=8)

        ttk.Radiobutton(
            of,
            text="Tesseract OCR — faster, requires Tesseract installed",
            variable=self.ocr_engine_var,
            value="tesseract",
        ).pack(anchor="w", pady=3)

        ttk.Radiobutton(
            of,
            text="EasyOCR — better Arabic accuracy, slower, no extra install needed",
            variable=self.ocr_engine_var,
            value="easyocr",
        ).pack(anchor="w", pady=3)

        # Table Mode
        tf = ttk.LabelFrame(tab, text=" Table Export Mode ", padding=10)
        tf.pack(fill="x", pady=8)

        for val, lbl in [
            ("separate", "Separate sheet per table"),
            ("per_page", "Group tables by page"),
            ("single",   "All tables in one sheet"),
        ]:
            ttk.Radiobutton(
                tf, text=lbl, variable=self.table_mode_var, value=val
            ).pack(anchor="w", pady=3)

        # Custom Patterns
        pf = ttk.LabelFrame(tab, text=" Custom Invoice Patterns (Advanced) ", padding=10)
        pf.pack(fill="both", expand=True, pady=8)

        ttk.Label(
            pf,
            text="One pattern per line.  Format:  field_name=regex_pattern",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(0, 4))

        self._patterns_box = ScrolledText(
            pf,
            height=7,
            font=("Consolas", 10),
            bg="#f8f9fa",
            relief="flat",
            wrap="word",
        )
        self._patterns_box.pack(fill="both", expand=True)
        self._patterns_box.insert(
            "end",
            "# Examples (lines starting with # are ignored):\n"
            "# po_number=(?:PO|order)\\s*[:\\-]?\\s*(\\S+)\n"
            "# customer=(?:Customer|client)\\s*[:\\-]?\\s*(.+)\n",
        )

    # ══════════════════════════════════════════════════════
    #  TAB 4 — About
    # ══════════════════════════════════════════════════════
    def _tab_about(self, nb):
        tab = ttk.Frame(nb, padding=20)
        nb.add(tab, text="   About   ")

        # App logo / title area
        logo_frame = tk.Frame(tab, bg="#1a3a6b", pady=20)
        logo_frame.pack(fill="x", pady=(0, 20))

        tk.Label(
            logo_frame,
            text=APP_NAME,
            font=("Segoe UI", 20, "bold"),
            fg="#ffffff",
            bg="#1a3a6b",
        ).pack()

        tk.Label(
            logo_frame,
            text=f"Version {APP_VERSION}",
            font=("Segoe UI", 11),
            fg="#aac4ff",
            bg="#1a3a6b",
        ).pack(pady=(4, 0))

        # Info cards
        cards_frame = ttk.Frame(tab)
        cards_frame.pack(fill="x", pady=10)
        cards_frame.columnconfigure(0, weight=1)
        cards_frame.columnconfigure(1, weight=1)

        # Card helper
        def make_card(parent, title, value, row, col, title_color="#1a3a6b", value_color="#000000"):
            card = tk.Frame(parent, bg="#f0f4ff", relief="flat", bd=1)
            card.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
            tk.Label(
                card,
                text=title,
                font=("Segoe UI", 9),
                fg="#666666",
                bg="#f0f4ff",
                pady=4,
            ).pack()
            tk.Label(
                card,
                text=value,
                font=("Segoe UI", 13, "bold"),
                fg=title_color,
                bg="#f0f4ff",
                pady=4,
            ).pack()

        make_card(cards_frame, "Created By",     CREATOR,          row=0, col=0, title_color="#1a3a6b")
        make_card(cards_frame, "Contact Number", CONTACT,          row=0, col=1, title_color="#c00000")
        make_card(cards_frame, "Version",        APP_VERSION,      row=1, col=0, title_color="#2e7d32")
        make_card(cards_frame, "Year",           YEAR,             row=1, col=1, title_color="#e65100")

        # Features
        features_frame = ttk.LabelFrame(tab, text=" Features ", padding=12)
        features_frame.pack(fill="x", pady=10)

        features = [
            "Extract data from text-based Arabic PDFs",
            "OCR support for scanned Arabic PDFs (Tesseract + EasyOCR)",
            "Table detection and extraction to Excel sheets",
            "Invoice and form field extraction using regex patterns",
            "Batch processing for multiple PDF files",
            "Full Arabic Unicode support — data is never changed",
            "All values saved as TEXT in Excel to prevent auto-formatting",
            "Right-to-Left (RTL) Excel sheets for Arabic content",
            "Auto-detect best extraction method for each PDF",
            "Desktop GUI with progress tracking",
        ]

        for feat in features:
            row = tk.Frame(features_frame, bg="white")
            row.pack(fill="x", pady=1)
            tk.Label(row, text="  ✔", font=("Segoe UI", 10, "bold"),
                     fg="#2e7d32", bg="white").pack(side="left")
            tk.Label(row, text=feat, font=("Segoe UI", 10),
                     fg="#333333", bg="white", anchor="w").pack(
                side="left", fill="x", expand=True, padx=6)

        # Copyright
        copy_frame = tk.Frame(tab, bg="#1a3a6b", pady=8)
        copy_frame.pack(fill="x", pady=(15, 0))

        tk.Label(
            copy_frame,
            text=f"© {YEAR}  {CREATOR}   |   All Rights Reserved   |   Contact: {CONTACT}",
            font=("Segoe UI", 10),
            fg="#ffffff",
            bg="#1a3a6b",
        ).pack()

    # ══════════════════════════════════════════════════════
    #  File Browsing
    # ══════════════════════════════════════════════════════
    def _browse_single(self):
        p = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if p:
            self.single_file_var.set(p)
            self._global_status.set("Selected: " + Path(p).name)

    def _browse_output(self):
        p = filedialog.askdirectory(title="Select Output Directory")
        if p:
            self.output_dir_var.set(p)

    # ══════════════════════════════════════════════════════
    #  PDF Analysis
    # ══════════════════════════════════════════════════════
    def _analyze_pdf(self):
        path = self.single_file_var.get().strip()
        if not self._check_file(path):
            return

        self._write_analysis("Analyzing PDF... please wait.")
        self._global_status.set("Analyzing...")

        def run():
            try:
                info = PDFAnalyzer.analyze(path)
                rec  = PDFAnalyzer.get_recommended_method(info)
                lines = [
                    "File       : " + Path(path).name,
                    "Size       : " + str(info["file_size_mb"]) + " MB",
                    "Pages      : " + str(info["page_count"]),
                    "Has Text   : " + ("Yes" if info["has_text"]   else "No"),
                    "Has Tables : " + ("Yes" if info["has_tables"] else "No"),
                    "Scanned    : " + ("Yes - OCR recommended" if info["is_scanned"] else "No"),
                    "",
                    "Recommended Method : " + rec.upper(),
                    "",
                    "Per-page breakdown:",
                ]
                for pg in info["pages_info"]:
                    lines.append(
                        "  Page " + str(pg["page_number"]).rjust(3)
                        + ":  Text=" + ("OK" if pg["has_text"]   else "--")
                        + "  Tables=" + ("OK" if pg["has_tables"] else "--")
                        + "  Images=" + ("OK" if pg["has_images"] else "--")
                        + "  chars=" + str(pg["text_length"])
                    )
                self.root.after(0, lambda: self._write_analysis("\n".join(lines)))
                self.root.after(0, lambda: self.method_var.set(rec))
                self.root.after(0, lambda: self._global_status.set(
                    "Analysis done - recommended: " + rec.upper()))

            except Exception as e:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self._write_analysis(
                    "Error:\n" + str(e) + "\n\n" + tb))

        threading.Thread(target=run, daemon=True).start()

    def _write_analysis(self, text):
        self._analysis_box.config(state="normal")
        self._analysis_box.delete("1.0", "end")
        self._analysis_box.insert("end", text)
        self._analysis_box.config(state="disabled")

    # ══════════════════════════════════════════════════════
    #  Single Extraction
    # ══════════════════════════════════════════════════════
    def _do_extract_single(self):
        path = self.single_file_var.get().strip()
        if not self._check_file(path):
            return
        if self.is_processing.get():
            messagebox.showinfo("Busy", "Already processing — please wait.")
            return

        self.is_processing.set(True)
        self._extract_btn.config(state="disabled")
        self._single_bar["value"] = 0
        self._single_status.set("Starting...")

        threading.Thread(target=self._run_single, args=(path,), daemon=True).start()

    def _run_single(self, pdf_path):
        try:
            out_dir = Path(self.output_dir_var.get())
            out_dir.mkdir(parents=True, exist_ok=True)
            xlsx = str(out_dir / (Path(pdf_path).stem + ".xlsx"))
            method = self.method_var.get()

            def cb(pct, msg):
                self.root.after(0, lambda p=pct, m=msg: self._upd_single(p, m))

            if method == "auto":
                self.root.after(0, lambda: self._single_status.set("Detecting PDF type..."))
                info   = PDFAnalyzer.analyze(pdf_path)
                method = PDFAnalyzer.get_recommended_method(info)

            if method == "ocr":
                ext   = OCRExtractor(engine=self.ocr_engine_var.get(), progress_callback=cb)
                stats = ext.extract_with_tables(pdf_path, xlsx)
            elif method == "table":
                ext   = TableExtractor(progress_callback=cb)
                stats = ext.extract(pdf_path, xlsx, mode=self.table_mode_var.get())
            elif method == "invoice":
                ext   = InvoiceExtractor(
                    custom_patterns=self._parse_patterns(), progress_callback=cb)
                stats = ext.extract(pdf_path, xlsx)
            else:
                ext   = TextExtractor(progress_callback=cb)
                stats = ext.extract(pdf_path, xlsx)

            self.root.after(0, lambda: self._single_done(xlsx, stats))

        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda err=str(e): self._single_error(err))

    def _upd_single(self, pct, msg):
        self._single_bar["value"] = pct
        self._single_status.set(msg)
        self._global_status.set(msg)

    def _single_done(self, xlsx, stats):
        self._single_bar["value"] = 100
        self._single_status.set("Done - " + Path(xlsx).name)
        self._global_status.set("Done: " + xlsx)
        self.is_processing.set(False)
        self._extract_btn.config(state="normal")
        stat_lines = "\n".join("  " + str(k) + ": " + str(v) for k, v in stats.items())
        if messagebox.askyesno(
            "Extraction Complete",
            "Extraction complete!\n\nFile:\n" + xlsx
            + "\n\nStats:\n" + stat_lines
            + "\n\nOpen output folder?"
        ):
            self._open_folder(Path(xlsx).parent)

    def _single_error(self, msg):
        self._single_bar["value"] = 0
        self._single_status.set("Error: " + msg)
        self._global_status.set("Error — see console for details")
        self.is_processing.set(False)
        self._extract_btn.config(state="normal")
        messagebox.showerror(
            "Error",
            "Extraction failed:\n\n" + msg + "\n\nSee terminal for full details."
        )

    # ══════════════════════════════════════════════════════
    #  Batch Processing
    # ══════════════════════════════════════════════════════
    def _batch_add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        added = 0
        for p in paths:
            if p not in self.pdf_files:
                self.pdf_files.append(p)
                sz = self._fmt_size(Path(p).stat().st_size)
                self._tree.insert("", "end", values=(Path(p).name, sz, "Pending"))
                added += 1
        self._batch_status.set(
            str(len(self.pdf_files)) + " file(s) ready  (+" + str(added) + " added)"
        )

    def _batch_add_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if not folder:
            return
        added = 0
        for f in sorted(Path(folder).glob("*.pdf")):
            p = str(f)
            if p not in self.pdf_files:
                self.pdf_files.append(p)
                sz = self._fmt_size(f.stat().st_size)
                self._tree.insert("", "end", values=(f.name, sz, "Pending"))
                added += 1
        self._batch_status.set(
            str(len(self.pdf_files)) + " file(s) ready  (+" + str(added) + " added)"
        )

    def _batch_clear(self):
        self.pdf_files.clear()
        for item in self._tree.get_children():
            self._tree.delete(item)
        self._batch_bar["value"] = 0
        self._batch_status.set("No files added")

    def _do_batch(self):
        if not self.pdf_files:
            messagebox.showwarning("No Files", "Please add PDF files first.")
            return
        if self.is_processing.get():
            messagebox.showinfo("Busy", "Already processing — please wait.")
            return
        self.is_processing.set(True)
        self._batch_btn.config(state="disabled")
        self._batch_bar["value"] = 0
        self._batch_status.set("Processing...")
        threading.Thread(target=self._run_batch, daemon=True).start()

    def _run_batch(self):
        try:
            items = self._tree.get_children()

            def cb(pct, msg):
                self.root.after(0, lambda p=pct, m=msg: self._upd_batch(p, m))

            proc = BatchProcessor(
                method=self.method_var.get(),
                ocr_engine=self.ocr_engine_var.get(),
                output_dir=self.output_dir_var.get(),
                progress_callback=cb,
            )
            results = proc.process_files(self.pdf_files)

            for idx, item in enumerate(items):
                if idx < len(results["files"]):
                    fr = results["files"][idx]
                    st = "Done" if fr["success"] else ("Failed: " + fr.get("error", "?")[:20])
                    vals = self._tree.item(item, "values")
                    self.root.after(
                        0,
                        lambda it=item, v=(vals[0], vals[1], st):
                            self._tree.item(it, values=v),
                    )

            self.root.after(0, lambda: self._batch_done(results))

        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            self.root.after(0, lambda err=str(e): self._batch_error(err))

    def _upd_batch(self, pct, msg):
        self._batch_bar["value"] = pct
        self._batch_status.set(msg)
        self._global_status.set(msg)

    def _batch_done(self, results):
        self._batch_bar["value"] = 100
        self._batch_status.set(
            "Done — Total: " + str(results["total"])
            + "  |  OK: " + str(results["success"])
            + "  |  Failed: " + str(results["failed"])
        )
        self._global_status.set("Batch complete")
        self.is_processing.set(False)
        self._batch_btn.config(state="normal")
        if messagebox.askyesno(
            "Batch Complete",
            "Batch done!\n\n"
            + "Total  : " + str(results["total"]) + "\n"
            + "Success: " + str(results["success"]) + "\n"
            + "Failed : " + str(results["failed"]) + "\n\n"
            + "Output folder:\n" + self.output_dir_var.get()
            + "\n\nOpen the output folder?"
        ):
            self._open_folder(self.output_dir_var.get())

    def _batch_error(self, msg):
        self._batch_bar["value"] = 0
        self._batch_status.set("Error: " + msg)
        self.is_processing.set(False)
        self._batch_btn.config(state="normal")
        messagebox.showerror("Batch Error", "Batch failed:\n\n" + msg)

    # ══════════════════════════════════════════════════════
    #  Helpers
    # ══════════════════════════════════════════════════════
    def _check_file(self, path):
        if not path:
            messagebox.showwarning("No File", "Please select a PDF file first.")
            return False
        if not os.path.isfile(path):
            messagebox.showerror("Not Found", "File not found:\n" + path)
            return False
        return True

    def _parse_patterns(self):
        text = self._patterns_box.get("1.0", "end").strip()
        out = {}
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if k and v:
                    out.setdefault(k, []).append(v)
        return out

    @staticmethod
    def _fmt_size(b):
        if b < 1024:
            return str(b) + " B"
        if b < 1024 * 1024:
            return str(round(b / 1024, 1)) + " KB"
        return str(round(b / (1024 * 1024), 2)) + " MB"

    @staticmethod
    def _open_folder(path):
        import subprocess
        try:
            if sys.platform == "win32":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass

    def run(self):
        self.root.mainloop()