"""
Advanced table extraction from PDFs.
"""

import pdfplumber
from utils.arabic_utils import clean_text, safe_sheet_name
from utils.config import LINE_TABLE_SETTINGS, TEXT_TABLE_SETTINGS
from core.excel_writer import ExcelWriter


class TableExtractor:
    """Extract tables from PDF files with multiple strategies."""

    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback

    def extract(self, pdf_path: str, excel_path: str, mode: str = "separate") -> dict:
        """
        Extract tables from PDF.

        mode:
          - 'separate': each table gets its own sheet
          - 'per_page': all tables on a page go into one sheet
          - 'single': all tables in one sheet
        """
        writer = ExcelWriter()
        stats = {"pages": 0, "tables": 0, "empty_pages": 0}
        all_tables = []

        with pdfplumber.open(pdf_path) as pdf:
            stats["pages"] = len(pdf.pages)

            for page_idx, page in enumerate(pdf.pages):
                page_num = page_idx + 1
                self._report_progress(page_num, stats["pages"])

                tables = self._extract_all_tables(page)

                if not tables:
                    stats["empty_pages"] += 1
                    continue

                for t_idx, table in enumerate(tables):
                    stats["tables"] += 1
                    cleaned = self._clean_table(table)

                    if mode == "separate":
                        name = safe_sheet_name(f"P{page_num}_T{t_idx + 1}")
                        writer.add_table_sheet(name, cleaned, has_header=True)
                    elif mode == "per_page":
                        all_tables.append((page_num, cleaned))
                    elif mode == "single":
                        all_tables.append(cleaned)

            # Write collected tables
            if mode == "per_page":
                pages_tables = {}
                for pn, tbl in all_tables:
                    pages_tables.setdefault(pn, []).append(tbl)
                for pn, tbls in pages_tables.items():
                    name = safe_sheet_name(f"Page_{pn}")
                    writer.add_multi_tables_sheet(name, tbls)

            elif mode == "single" and all_tables:
                writer.add_multi_tables_sheet("All_Tables", all_tables)

        writer.save(excel_path)
        return stats

    def _extract_all_tables(self, page) -> list:
        """Extract tables using multiple strategies."""
        best_tables = []

        for settings in (LINE_TABLE_SETTINGS, TEXT_TABLE_SETTINGS):
            try:
                tables = page.extract_tables(table_settings=settings) or []
                valid = [t for t in tables if self._is_valid(t)]
                if len(valid) > len(best_tables):
                    best_tables = valid
            except Exception:
                continue

        return best_tables

    def _clean_table(self, table: list) -> list:
        """Clean all cells in a table."""
        cleaned = []
        max_cols = max((len(row) for row in table if row), default=0)
        for row in table:
            if row is None:
                cleaned.append([""] * max_cols)
                continue
            cleaned_row = [clean_text(str(cell)) if cell else "" for cell in row]
            # Pad
            while len(cleaned_row) < max_cols:
                cleaned_row.append("")
            cleaned.append(cleaned_row)
        return cleaned

    @staticmethod
    def _is_valid(table: list) -> bool:
        """Validate table has data."""
        if not table or len(table) < 1:
            return False
        cells_with_data = sum(
            1 for row in table if row
            for cell in row if cell and str(cell).strip()
        )
        return cells_with_data >= 2

    def _report_progress(self, current: int, total: int):
        if self.progress_callback:
            pct = int((current / total) * 100)
            self.progress_callback(pct, f"Extracting tables: page {current}/{total}")