#!/usr/bin/env python3
"""
Arabic PDF to Excel Extractor
Supports: Text PDFs, Scanned PDFs (OCR), Tables, Invoices/Forms, Batch Processing

Usage:
  GUI Mode:    python main.py
  CLI Mode:    python main.py --cli --input file.pdf --output result.xlsx --method auto
  Batch Mode:  python main.py --cli --batch --input ./pdfs/ --output ./output/ --method auto
"""

import argparse
import sys
import os
from pathlib import Path


def run_gui():
    """Launch the GUI application."""
    from gui import PDFExtractorApp
    app = PDFExtractorApp()
    app.run()


def run_cli(args):
    """Run in CLI mode."""
    from core.text_extractor import TextExtractor
    from core.ocr_extractor import OCRExtractor
    from core.table_extractor import TableExtractor
    from core.invoice_extractor import InvoiceExtractor
    from core.batch_processor import BatchProcessor
    from utils.pdf_analyzer import PDFAnalyzer

    def progress(pct, msg):
        print(f"\r  [{pct:3d}%] {msg}", end="", flush=True)

    if args.batch:
        # Batch mode
        input_dir = args.input
        output_dir = args.output or str(Path.home() / "Desktop" / "PDF_Extracted")

        print(f"\n📁 Batch Processing: {input_dir}")
        print(f"📂 Output Directory: {output_dir}")
        print(f"🔧 Method: {args.method}")
        print(f"🔍 OCR Engine: {args.ocr_engine}\n")

        processor = BatchProcessor(
            method=args.method,
            ocr_engine=args.ocr_engine,
            output_dir=output_dir,
            progress_callback=progress,
        )
        results = processor.process_directory(input_dir, recursive=args.recursive)

        print(f"\n\n{'='*50}")
        print(f"✅ Batch Complete!")
        print(f"   Total:   {results['total']}")
        print(f"   Success: {results['success']}")
        print(f"   Failed:  {results['failed']}")

        for f in results["files"]:
            status = "✅" if f["success"] else "❌"
            print(f"   {status} {Path(f['input']).name} → {f.get('method', '?')} — {f.get('error', 'OK')}")

    else:
        # Single file mode
        pdf_path = args.input
        if not os.path.isfile(pdf_path):
            print(f"❌ File not found: {pdf_path}")
            sys.exit(1)

        output_dir = args.output or str(Path.home() / "Desktop" / "PDF_Extracted")
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        pdf_name = Path(pdf_path).stem
        excel_path = os.path.join(output_dir, f"{pdf_name}.xlsx")

        method = args.method

        print(f"\n📄 Input:  {pdf_path}")
        print(f"📊 Output: {excel_path}")

        # Auto-detect
        if method == "auto":
            print("🔍 Analyzing PDF...")
            analysis = PDFAnalyzer.analyze(pdf_path)
            method = PDFAnalyzer.get_recommended_method(analysis)
            print(f"   Pages: {analysis['page_count']}, Scanned: {analysis['is_scanned']}")
            print(f"   Tables: {analysis['has_tables']}, Text: {analysis['has_text']}")

        print(f"🔧 Method: {method}")
        print()

        if method == "ocr":
            extractor = OCRExtractor(engine=args.ocr_engine, progress_callback=progress)
            stats = extractor.extract_with_tables(pdf_path, excel_path)
        elif method == "table":
            extractor = TableExtractor(progress_callback=progress)
            stats = extractor.extract(pdf_path, excel_path, mode="separate")
        elif method == "invoice":
            extractor = InvoiceExtractor(progress_callback=progress)
            stats = extractor.extract(pdf_path, excel_path)
        else:
            extractor = TextExtractor(progress_callback=progress)
            stats = extractor.extract(pdf_path, excel_path)

        print(f"\n\n✅ Done! Saved to: {excel_path}")
        print(f"   Stats: {stats}")


def main():
    parser = argparse.ArgumentParser(
        description="Arabic PDF to Excel Extractor",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py                                          # Launch GUI
  python main.py --cli -i invoice.pdf -m invoice          # Extract invoice
  python main.py --cli -i scan.pdf -m ocr --ocr easyocr   # OCR with EasyOCR
  python main.py --cli --batch -i ./pdfs/ -m auto          # Batch process
        """,
    )

    parser.add_argument("--cli", action="store_true", help="Run in CLI mode (default: GUI)")
    parser.add_argument("-i", "--input", help="Input PDF file or directory (for batch)")
    parser.add_argument("-o", "--output", help="Output Excel file or directory")
    parser.add_argument(
        "-m", "--method",
        choices=["auto", "text", "ocr", "table", "invoice"],
        default="auto",
        help="Extraction method (default: auto)",
    )
    parser.add_argument(
        "--ocr",
        dest="ocr_engine",
        choices=["tesseract", "easyocr"],
        default="tesseract",
        help="OCR engine (default: tesseract)",
    )
    parser.add_argument("--batch", action="store_true", help="Batch process all PDFs in input directory")
    parser.add_argument("--recursive", action="store_true", help="Include subdirectories in batch mode")

    args = parser.parse_args()

    if args.cli:
        if not args.input:
            parser.error("--cli mode requires --input")
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()