"""
Arabic text processing utilities.
Fixes Arabic text that comes out as reversed/broken letters.
"""

import re
import unicodedata

# ── Try to import Arabic fixers ──
try:
    import arabic_reshaper
    HAS_RESHAPER = True
except ImportError:
    HAS_RESHAPER = False
    print("WARNING: arabic_reshaper not installed. Run: pip install arabic-reshaper")

try:
    from bidi.algorithm import get_display
    HAS_BIDI = True
except ImportError:
    HAS_BIDI = False
    print("WARNING: python-bidi not installed. Run: pip install python-bidi")


def fix_arabic_text(text: str) -> str:
    """
    Fix Arabic text extracted from PDFs.
    Handles:
      - Reversed letters
      - Disconnected letters
      - Wrong display order (RTL)
      - Mixed Arabic/English text
    """
    if not text:
        return ""

    if not HAS_RESHAPER or not HAS_BIDI:
        return text  # Return as-is if libraries not available

    lines = text.splitlines()
    fixed_lines = []

    for line in lines:
        if not line.strip():
            fixed_lines.append(line)
            continue

        if _contains_arabic(line):
            try:
                # Step 1: Reshape Arabic letters (connect them properly)
                reshaped = arabic_reshaper.reshape(line)
                # Step 2: Fix display order (RTL)
                display  = get_display(reshaped)
                fixed_lines.append(display)
            except Exception:
                fixed_lines.append(line)
        else:
            fixed_lines.append(line)

    return "\n".join(fixed_lines)


def fix_arabic_cell(value: str) -> str:
    """
    Fix a single cell value (Arabic text in a table cell).
    This is the main function used for Excel cell values.
    """
    if not value or not isinstance(value, str):
        return value if value else ""

    value = value.strip()

    if not _contains_arabic(value):
        return value  # English/numbers — no fix needed

    if not HAS_RESHAPER or not HAS_BIDI:
        return value

    try:
        reshaped = arabic_reshaper.reshape(value)
        display  = get_display(reshaped)
        return display
    except Exception:
        return value


def fix_arabic_table(table: list) -> list:
    """
    Fix all Arabic text in a table (list of lists).
    Returns the fixed table.
    """
    if not table:
        return table

    fixed_table = []
    for row in table:
        if not row:
            fixed_table.append(row)
            continue
        fixed_row = []
        for cell in row:
            if cell and isinstance(cell, str):
                fixed_row.append(fix_arabic_cell(cell))
            else:
                fixed_row.append(cell if cell is not None else "")
        fixed_table.append(fixed_row)

    return fixed_table


def fix_arabic_dict(data: dict) -> dict:
    """Fix Arabic text in a dictionary (key-value pairs)."""
    fixed = {}
    for key, value in data.items():
        fixed_key   = fix_arabic_cell(str(key))   if key   else key
        fixed_value = fix_arabic_cell(str(value)) if value else value
        fixed[fixed_key] = fixed_value
    return fixed


def _contains_arabic(text: str) -> bool:
    """Check if text contains Arabic characters."""
    if not text:
        return False
    for char in text:
        code = ord(char)
        if (
            0x0600 <= code <= 0x06FF or   # Arabic
            0x0750 <= code <= 0x077F or   # Arabic Supplement
            0xFB50 <= code <= 0xFDFF or   # Arabic Presentation Forms-A
            0xFE70 <= code <= 0xFEFF      # Arabic Presentation Forms-B
        ):
            return True
    return False


def is_arabic(text: str) -> bool:
    """Public alias for _contains_arabic."""
    return _contains_arabic(text)


def clean_text(text: str) -> str:
    """Clean text while preserving Arabic characters."""
    if not text:
        return ""
    text = text.replace("\u00A0", " ")   # non-breaking space
    text = text.replace("\ufeff", "")    # BOM
    text = text.replace("\u200b", "")    # zero-width space
    text = text.replace("\u200c", "")    # zero-width non-joiner
    text = text.replace("\u200d", "")    # zero-width joiner
    text = text.replace("\u200e", "")    # LTR mark
    text = text.replace("\u200f", "")    # RTL mark
    text = re.sub(r" {3,}", "  ", text)
    return text.strip()


def normalize_arabic(text: str) -> str:
    """Normalize Arabic text for consistent matching."""
    if not text:
        return ""
    text = re.sub(r"[إأآا]", "ا", text)      # Normalize alef
    text = re.sub(r"ة", "ه", text)            # Normalize taa marbuta
    text = re.sub(r"ى", "ي", text)            # Normalize yaa
    text = re.sub(r"[\u064B-\u065F\u0670]", "", text)  # Remove diacritics
    text = text.replace("\u0640", "")         # Remove tatweel
    return text


def fix_arabic_display(text: str) -> str:
    """Fix common OCR issues with Arabic text then reshape."""
    if not text:
        return ""
    replacements = {
        "لأ": "لأ",
        "لإ": "لإ",
        "لآ": "لآ",
    }
    for wrong, correct in replacements.items():
        text = text.replace(wrong, correct)
    return fix_arabic_text(text)


def safe_sheet_name(name: str) -> str:
    """Create a valid Excel sheet name."""
    name = re.sub(r'[:\\/*?\[\]]', "_", name)
    name = name[:31] if len(name) > 31 else name
    return name if name else "Sheet"


def detect_text_direction(text: str) -> str:
    """Detect if text is primarily RTL or LTR."""
    return "rtl" if is_arabic(text) else "ltr"