"""
Batch processing for multiple PDF files.
"""

import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from utils.pdf_analyzer import PDFAnalyzer
from utils.config import SUPPORTED_EXTENSIONS, DEFAULT_OUTPUT_DIR
from core.text_extractor import TextExtractor
from core.ocr_extractor import OCRExtractor
from core.table_extractor import TableExtractor
from core.invoice_extractor import InvoiceExtractor


class BatchProcessor:
    """Process multiple PDF files in batch."""

    def __init__(
        self,
        method: str = "auto",
        ocr_engine: str = "tesseract",
        output_dir: str = None,
        max_workers: int = 2,
        progress_callback=None,
    ):
        """
        method: 'auto', 'text', 'ocr', 'table', 'invoice'
        ocr_engine: 'tesseract' or 'easyocr'
        """
        self.method = method
        self.ocr_engine = ocr_engine
        self.output_dir = Path(output_dir) if output_dir else DEFAULT_OUTPUT_DIR
        self.max_workers = max_workers
        self.progress_callback = progress_callback

    def process_files(self, pdf_paths: list) -> dict:
        """Process a list of PDF files."""
        self.output_dir.mkdir(parents=True, exist_ok=True)

        results = {
            "total": len(pdf_paths),
            "success": 0,
            "failed": 0,
            "files": [],
        }

        for idx, pdf_path in enumerate(pdf_paths, start=1):
            self._report_progress(
                int((idx / len(pdf_paths)) * 100),
                f"Processing file {idx}/{len(pdf_paths)}: {Path(pdf_path).name}",
            )

            result = self._process_single(pdf_path)
            results["files"].append(result)

            if result["success"]:
                results["success"] += 1
            else:
                results["failed"] += 1

        return results

    def process_directory(self, directory: str, recursive: bool = False) -> dict:
        """Process all PDFs in a directory."""
        dir_path = Path(directory)
        if recursive:
            pdf_files = [
                str(f) for f in dir_path.rglob("*")
                if f.suffix.lower() in SUPPORTED_EXTENSIONS
            ]
        else:
            pdf_files = [
                str(f) for f in dir_path.iterdir()
                if f.suffix.lower() in SUPPORTED_EXTENSIONS
            ]

        if not pdf_files:
            return {"total": 0, "success": 0, "failed": 0, "files": [], "message": "No PDF files found."}

        return self.process_files(pdf_files)

    def _process_single(self, pdf_path: str) -> dict:
        """Process a single PDF file."""
        result = {
            "input": pdf_path,
            "output": "",
            "success": False,
            "method": "",
            "stats": {},
            "error": "",
        }

        try:
            pdf_name = Path(pdf_path).stem
            excel_path = str(self.output_dir / f"{pdf_name}.xlsx")
            result["output"] = excel_path

            # Determine method
            method = self.method
            if method == "auto":
                analysis = PDFAnalyzer.analyze(pdf_path)
                method = PDFAnalyzer.get_recommended_method(analysis)

            result["method"] = method

            # Extract
            if method == "ocr":
                extractor = OCRExtractor(engine=self.ocr_engine)
                stats = extractor.extract_with_tables(pdf_path, excel_path)
            elif method == "table":
                extractor = TableExtractor()
                stats = extractor.extract(pdf_path, excel_path, mode="separate")
            elif method == "invoice":
                extractor = InvoiceExtractor()
                stats = extractor.extract(pdf_path, excel_path)
            else:
                extractor = TextExtractor()
                stats = extractor.extract(pdf_path, excel_path)

            result["stats"] = stats
            result["success"] = True

        except Exception as e:
            result["error"] = str(e)
            result["success"] = False

        return result

    def _report_progress(self, pct: int, message: str):
        if self.progress_callback:
            self.progress_callback(pct, message)