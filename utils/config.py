"""
Configuration constants for the Arabic PDF Extractor.
"""

import os
import sys
from pathlib import Path

# ---------- Tesseract path ----------
if sys.platform == "win32":
    TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
else:
    TESSERACT_CMD = "tesseract"

# ---------- Poppler path (Windows) ----------
if sys.platform == "win32":
    POPPLER_PATH = r"C:\poppler\Library\bin"
else:
    POPPLER_PATH = None

# ---------- OCR settings ----------
OCR_LANGUAGES = "ara+eng"
OCR_DPI = 300
OCR_PSM = 6  # Assume uniform block of text

# ---------- Table extraction ----------
LINE_TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 4,
    "intersection_tolerance": 6,
}

TEXT_TABLE_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "snap_tolerance": 4,
    "intersection_tolerance": 6,
    "min_words_vertical": 1,
    "min_words_horizontal": 1,
}

# ---------- Excel ----------
EXCEL_FONT_NAME = "Arial"
EXCEL_FONT_SIZE = 11
MAX_COLUMN_WIDTH = 60
SHEET_NAME_MAX_LEN = 31

# ---------- Invoice / form patterns ----------
INVOICE_PATTERNS = {
    "invoice_number": [
        r"(?:رقم\s*الفاتورة|فاتورة\s*رقم|Invoice\s*(?:No\.?|Number|#))\s*[:\-]?\s*(\S+)",
    ],
    "date": [
        r"(?:التاريخ|تاريخ|Date)\s*[:\-]?\s*(\d{1,4}[\-/\.]\d{1,2}[\-/\.]\d{1,4})",
    ],
    "total": [
        r"(?:الإجمالي|المجموع|المبلغ\s*الإجمالي|Total|Grand\s*Total|Amount\s*Due)\s*[:\-]?\s*([\d,\.]+)",
    ],
    "vat": [
        r"(?:ضريبة|الضريبة|VAT|Tax)\s*[:\-]?\s*([\d,\.]+)",
    ],
    "company_name": [
        r"(?:اسم\s*الشركة|الشركة|Company|Vendor|Supplier)\s*[:\-]?\s*(.+)",
    ],
    "vat_number": [
        r"(?:الرقم\s*الضريبي|رقم\s*ضريبي|VAT\s*(?:No\.?|Number|ID|TIN))\s*[:\-]?\s*(\d{10,20})",
    ],
    "phone": [
        r"(?:هاتف|جوال|تلفون|Phone|Tel|Mobile)\s*[:\-]?\s*([\d\+\-\s\(\)]{7,20})",
    ],
    "email": [
        r"([\w\.\-]+@[\w\.\-]+\.\w{2,})",
    ],
    "subtotal": [
        r"(?:المجموع\s*الفرعي|المجموع\s*قبل|Subtotal|Sub\s*Total)\s*[:\-]?\s*([\d,\.]+)",
    ],
}

# ---------- Supported file types ----------
SUPPORTED_EXTENSIONS = {".pdf"}

# ---------- Output directory ----------
DEFAULT_OUTPUT_DIR = Path.home() / "Desktop" / "PDF_Extracted"