"""
Diagnostic script — checks all dependencies.
Run: python check_env.py
"""

import sys
import importlib

print(f"Python version: {sys.version}\n")

packages = [
    ("pdfplumber",              "pdfplumber"),
    ("openpyxl",                "openpyxl"),
    ("Pillow",                  "PIL"),
    ("pytesseract",             "pytesseract"),
    ("pdf2image",               "pdf2image"),
    ("easyocr",                 "easyocr"),
    ("numpy",                   "numpy"),
    ("opencv-python-headless",  "cv2"),
    ("ttkbootstrap",            "ttkbootstrap"),
]

ok = []
missing = []

for pkg_name, import_name in packages:
    try:
        mod = importlib.import_module(import_name)
        version = getattr(mod, "__version__", "unknown version")
        print(f"  ✅ {pkg_name:<30} {version}")
        ok.append(pkg_name)
    except ImportError:
        print(f"  ❌ {pkg_name:<30} NOT INSTALLED")
        missing.append(pkg_name)

# Check ttkbootstrap sub-modules
print()
sub_checks = [
    ("ttkbootstrap.constants",  "constants"),
    ("tkinter.scrolledtext",    "ScrolledText (built-in — always available)"),
]
for mod_path, label in sub_checks:
    try:
        importlib.import_module(mod_path)
        print(f"  ✅ {mod_path} ({label})")
    except ImportError:
        print(f"  ❌ {mod_path} — {label} NOT available")

# Tesseract executable
print()
try:
    import pytesseract
    version = pytesseract.get_tesseract_version()
    print(f"  ✅ Tesseract executable found — version {version}")
except Exception as e:
    print(f"  ⚠️  Tesseract executable not found ({e})")
    print("     Download from: https://github.com/UB-Mannheim/tesseract/wiki")

print()
print("=" * 55)
if missing:
    print(f"Install missing packages:")
    print(f"  pip install {' '.join(missing)}")
else:
    print("✅ All packages installed. Ready to run!")
print("=" * 55)