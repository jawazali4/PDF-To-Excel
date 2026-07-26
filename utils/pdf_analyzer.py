"""
Analyze PDF to determine type (text-based or scanned/image).
"""

import pdfplumber
from pathlib import Path


class PDFAnalyzer:
    """Analyze PDF files to determine extraction strategy."""

    @staticmethod
    def analyze(pdf_path: str) -> dict:
        """
        Analyze a PDF file.

        Returns dict with:
          - page_count: int
          - is_scanned: bool
          - has_tables: bool
          - has_text: bool
          - pages_info: list of per-page info
        """
        result = {
            "page_count": 0,
            "is_scanned": False,
            "has_tables": False,
            "has_text": False,
            "pages_info": [],
            "file_size_mb": round(Path(pdf_path).stat().st_size / (1024 * 1024), 2),
        }

        with pdfplumber.open(pdf_path) as pdf:
            result["page_count"] = len(pdf.pages)

            text_pages = 0
            image_pages = 0
            table_pages = 0

            for i, page in enumerate(pdf.pages):
                page_info = {
                    "page_number": i + 1,
                    "has_text": False,
                    "has_tables": False,
                    "has_images": False,
                    "text_length": 0,
                }

                # Check text
                text = page.extract_text() or ""
                text = text.strip()
                page_info["text_length"] = len(text)

                if len(text) > 20:
                    page_info["has_text"] = True
                    text_pages += 1

                # Check tables
                tables = page.extract_tables() or []
                non_empty = [
                    t for t in tables
                    if t and any(
                        any(cell and str(cell).strip() for cell in row)
                        for row in t if row
                    )
                ]
                if non_empty:
                    page_info["has_tables"] = True
                    table_pages += 1

                # Check images
                if hasattr(page, "images") and page.images:
                    page_info["has_images"] = True
                    if not page_info["has_text"]:
                        image_pages += 1

                result["pages_info"].append(page_info)

            result["has_text"] = text_pages > 0
            result["has_tables"] = table_pages > 0

            # If most pages are image-only, the PDF is scanned
            if result["page_count"] > 0:
                result["is_scanned"] = (image_pages / result["page_count"]) > 0.5

        return result

    @staticmethod
    def get_recommended_method(analysis: dict) -> str:
        """Get recommended extraction method."""
        if analysis["is_scanned"]:
            return "ocr"
        if analysis["has_tables"]:
            return "table"
        if analysis["has_text"]:
            return "text"
        return "ocr"