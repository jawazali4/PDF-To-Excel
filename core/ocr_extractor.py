"""
OCR-based extraction for scanned Arabic PDFs.
Supports both Tesseract and EasyOCR.
"""

import os
import tempfile
from pathlib import Path

from utils.arabic_utils import clean_text, fix_arabic_display, safe_sheet_name
from utils.config import OCR_LANGUAGES, OCR_DPI, OCR_PSM, TESSERACT_CMD, POPPLER_PATH
from core.excel_writer import ExcelWriter


class OCRExtractor:
    """Extract text from scanned PDFs using OCR."""

    def __init__(self, engine: str = "tesseract", progress_callback=None):
        """
        engine: 'tesseract' or 'easyocr'
        """
        self.engine = engine.lower()
        self.progress_callback = progress_callback
        self._easyocr_reader = None

    def extract(self, pdf_path: str, excel_path: str) -> dict:
        """Extract scanned PDF to Excel using OCR."""
        writer = ExcelWriter()
        stats = {"pages": 0, "text_pages": 0, "engine": self.engine}

        images = self._pdf_to_images(pdf_path)
        stats["pages"] = len(images)

        for page_idx, img in enumerate(images):
            page_num = page_idx + 1
            self._report_progress(page_num, stats["pages"])

            if self.engine == "easyocr":
                text = self._ocr_easyocr(img)
            else:
                text = self._ocr_tesseract(img)

            text = fix_arabic_display(clean_text(text))

            if text.strip():
                stats["text_pages"] += 1
                lines = text.splitlines()
                sheet_name = safe_sheet_name(f"Page_{page_num}")
                writer.add_text_sheet(sheet_name, lines)

        writer.save(excel_path)
        return stats

    def extract_with_tables(self, pdf_path: str, excel_path: str) -> dict:
        """Extract scanned PDF with table detection using OCR."""
        writer = ExcelWriter()
        stats = {"pages": 0, "tables": 0, "text_pages": 0, "engine": self.engine}

        images = self._pdf_to_images(pdf_path)
        stats["pages"] = len(images)

        for page_idx, img in enumerate(images):
            page_num = page_idx + 1
            self._report_progress(page_num, stats["pages"])

            # Try table detection with OpenCV
            table_data = self._detect_table_from_image(img)

            if table_data:
                stats["tables"] += 1
                sheet_name = safe_sheet_name(f"P{page_num}_T1")
                writer.add_table_sheet(sheet_name, table_data, has_header=True)
            else:
                if self.engine == "easyocr":
                    text = self._ocr_easyocr(img)
                else:
                    text = self._ocr_tesseract(img)

                text = fix_arabic_display(clean_text(text))
                if text.strip():
                    stats["text_pages"] += 1
                    lines = text.splitlines()
                    sheet_name = safe_sheet_name(f"Page_{page_num}")
                    writer.add_text_sheet(sheet_name, lines)

        writer.save(excel_path)
        return stats

    def _pdf_to_images(self, pdf_path: str) -> list:
        """Convert PDF pages to images."""
        from pdf2image import convert_from_path

        kwargs = {"dpi": OCR_DPI}
        if POPPLER_PATH and os.path.isdir(POPPLER_PATH):
            kwargs["poppler_path"] = POPPLER_PATH

        return convert_from_path(pdf_path, **kwargs)

    def _ocr_tesseract(self, image) -> str:
        """OCR using Tesseract."""
        import pytesseract

        if os.path.isfile(TESSERACT_CMD):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

        config = f"--psm {OCR_PSM} --oem 3"
        text = pytesseract.image_to_string(image, lang=OCR_LANGUAGES, config=config)
        return text or ""

    def _ocr_easyocr(self, image) -> str:
        """OCR using EasyOCR."""
        import easyocr
        import numpy as np

        if self._easyocr_reader is None:
            self._easyocr_reader = easyocr.Reader(["ar", "en"], gpu=False)

        img_array = np.array(image)
        results = self._easyocr_reader.readtext(img_array, detail=1, paragraph=True)

        # Sort by vertical position (top to bottom)
        results.sort(key=lambda r: r[0][0][1] if r[0] else 0)

        lines = [r[1] for r in results if r[1].strip()]
        return "\n".join(lines)

    def _detect_table_from_image(self, image) -> list:
        """Detect and extract table from image using OpenCV."""
        try:
            import cv2
            import numpy as np

            img_array = np.array(image)
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

            # Detect horizontal lines
            h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
            h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel, iterations=2)

            # Detect vertical lines
            v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
            v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel, iterations=2)

            # Combine
            table_mask = cv2.add(h_lines, v_lines)

            contours, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

            if len(contours) < 4:
                return []

            # Find cells
            cells = []
            for cnt in contours:
                x, y, w, h = cv2.boundingRect(cnt)
                if 20 < w < img_array.shape[1] * 0.95 and 10 < h < img_array.shape[0] * 0.5:
                    cells.append((x, y, w, h))

            if len(cells) < 4:
                return []

            # Sort cells into rows
            cells.sort(key=lambda c: (c[1], c[0]))

            rows = []
            current_row = [cells[0]]
            for cell in cells[1:]:
                if abs(cell[1] - current_row[0][1]) < 15:
                    current_row.append(cell)
                else:
                    current_row.sort(key=lambda c: c[0])
                    rows.append(current_row)
                    current_row = [cell]
            if current_row:
                current_row.sort(key=lambda c: c[0])
                rows.append(current_row)

            # OCR each cell
            table_data = []
            for row_cells in rows:
                row_texts = []
                for (x, y, w, h) in row_cells:
                    cell_img = image.crop((x, y, x + w, y + h))
                    if self.engine == "easyocr":
                        text = self._ocr_easyocr(cell_img)
                    else:
                        text = self._ocr_tesseract(cell_img)
                    row_texts.append(clean_text(text))
                table_data.append(row_texts)

            return table_data if len(table_data) >= 2 else []

        except Exception:
            return []

    def _report_progress(self, current: int, total: int):
        """Report progress."""
        if self.progress_callback:
            pct = int((current / total) * 100)
            self.progress_callback(pct, f"OCR page {current}/{total}")