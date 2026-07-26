"""
Invoice and form field extraction.
Uses regex patterns to find structured data like invoice numbers, dates, totals, etc.
"""

import re
import pdfplumber
from utils.arabic_utils import clean_text, safe_sheet_name, normalize_arabic
from utils.config import INVOICE_PATTERNS
from core.excel_writer import ExcelWriter
from core.table_extractor import TableExtractor


class InvoiceExtractor:
    """Extract structured fields from invoices and forms."""

    def __init__(self, custom_patterns: dict = None, progress_callback=None):
        self.patterns = {**INVOICE_PATTERNS}
        if custom_patterns:
            self.patterns.update(custom_patterns)
        self.progress_callback = progress_callback
        self.table_extractor = TableExtractor(progress_callback)

    def extract(self, pdf_path: str, excel_path: str) -> dict:
        """Extract invoice fields and tables from PDF."""
        writer = ExcelWriter()
        stats = {"pages": 0, "fields_found": 0, "tables": 0}

        all_text = ""
        page_texts = []

        with pdfplumber.open(pdf_path) as pdf:
            stats["pages"] = len(pdf.pages)

            for page_idx, page in enumerate(pdf.pages):
                page_num = page_idx + 1
                self._report_progress(page_num, stats["pages"])

                text = page.extract_text(layout=True) or ""
                text = clean_text(text)
                page_texts.append(text)
                all_text += "\n" + text

        # Extract fields
        fields = self._extract_fields(all_text)
        stats["fields_found"] = sum(1 for v in fields.values() if v)

        # Add fields sheet
        if fields:
            display_fields = {}
            field_labels = {
                "invoice_number": "رقم الفاتورة / Invoice Number",
                "date": "التاريخ / Date",
                "total": "الإجمالي / Total",
                "vat": "الضريبة / VAT",
                "company_name": "الشركة / Company",
                "vat_number": "الرقم الضريبي / VAT Number",
                "phone": "الهاتف / Phone",
                "email": "البريد الإلكتروني / Email",
                "subtotal": "المجموع الفرعي / Subtotal",
            }
            for key, value in fields.items():
                label = field_labels.get(key, key)
                display_fields[label] = value if value else "—"

            writer.add_key_value_sheet("Invoice_Fields", display_fields)

        # Extract tables
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                page_num = page_idx + 1
                tables = self.table_extractor._extract_all_tables(page)
                for t_idx, table in enumerate(tables):
                    stats["tables"] += 1
                    cleaned = self.table_extractor._clean_table(table)
                    name = safe_sheet_name(f"P{page_num}_Table{t_idx + 1}")
                    writer.add_table_sheet(name, cleaned, has_header=True)

        # Add raw text
        all_lines = []
        for p_idx, text in enumerate(page_texts):
            if text.strip():
                all_lines.append(f"=== Page {p_idx + 1} ===")
                all_lines.extend(text.splitlines())
                all_lines.append("")

        if all_lines:
            writer.add_text_sheet("Raw_Text", all_lines)

        writer.save(excel_path)
        return stats

    def extract_from_text(self, text: str) -> dict:
        """Extract fields from text (for OCR integration)."""
        return self._extract_fields(text)

    def _extract_fields(self, text: str) -> dict:
        """Extract all fields using patterns."""
        fields = {}
        normalized = normalize_arabic(text)

        for field_name, patterns in self.patterns.items():
            value = None
            for pattern in patterns:
                try:
                    # Search in original text
                    match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
                    if not match:
                        # Search in normalized text
                        match = re.search(pattern, normalized, re.IGNORECASE | re.MULTILINE)
                    if match:
                        value = match.group(1).strip()
                        break
                except Exception:
                    continue
            fields[field_name] = value

        return fields

    def _report_progress(self, current: int, total: int):
        if self.progress_callback:
            pct = int((current / total) * 100)
            self.progress_callback(pct, f"Analyzing invoice: page {current}/{total}")