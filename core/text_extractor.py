"""
Text-based PDF extraction using pdfplumber.
Applies Arabic text fix before writing to Excel.
"""

import pdfplumber
from utils.arabic_utils import clean_text, fix_arabic_text, safe_sheet_name
from utils.config import LINE_TABLE_SETTINGS, TEXT_TABLE_SETTINGS
from core.excel_writer import ExcelWriter


class TextExtractor:
    """Extract data from text-based PDFs."""

    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

    def extract(self, pdf_path: str, excel_path: str) -> dict:
        writer = ExcelWriter()
        stats  = {"pages": 0, "tables": 0, "text_pages": 0}

        with pdfplumber.open(pdf_path) as pdf:
            stats["pages"] = len(pdf.pages)

            for page_idx, page in enumerate(pdf.pages):
                page_num = page_idx + 1
                self._report(page_num, stats["pages"])

                tables = self._extract_tables(page)

                if tables:
                    for t_idx, table in enumerate(tables):
                        stats["tables"] += 1
                        name = safe_sheet_name(f"P{page_num}_T{t_idx + 1}")
                        # Arabic fix is applied inside ExcelWriter.add_table_sheet
                        writer.add_table_sheet(name, table, has_header=True)
                else:
                    text = page.extract_text(layout=True) or ""
                    text = clean_text(text)
                    if text:
                        stats["text_pages"] += 1
                        # Fix Arabic text before splitting into lines
                        fixed = fix_arabic_text(text)
                        lines = fixed.splitlines()
                        name  = safe_sheet_name(f"Page_{page_num}")
                        writer.add_text_sheet(name, lines)

        writer.save(excel_path)
        return stats

    def _extract_tables(self, page) -> list:
        for settings in (LINE_TABLE_SETTINGS, TEXT_TABLE_SETTINGS):
            try:
                tables = page.extract_tables(table_settings=settings) or []
                valid  = [t for t in tables if self._is_valid(t)]
                if valid:
                    return valid
            except Exception:
                continue
        return []

    @staticmethod
    def _is_valid(table: list) -> bool:
        if not table:
            return False
        return sum(
            1 for row in table if row
            for cell in row if cell and str(cell).strip()
        ) >= 2

    def _report(self, current: int, total: int):
        if self.progress_callback:
            self.progress_callback(
                int((current / total) * 100),
                f"Processing page {current}/{total}"
            )