"""
utils/arabic.py — Pure Arabic text utilities (no GUI, no Excel).
"""

import re
import pandas as pd


def extract_code(val) -> str:
    """Extract numeric office code from '3000-Alexandria...' → '3000'."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    m = re.match(r"^(\d+)", str(val).strip())
    return m.group(1) if m else str(val).strip()


def row_notes(total_items: float) -> str:
    """Return shipment notes based on item count."""
    return "على المكشوف - قابل للكسر" if total_items == 1 else "قابل للكسر"


def tafqeet(n: float) -> str:
    """Convert a number to Arabic words (تفقيط)."""
    ones = [
        "", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة",
        "سبعة", "ثمانية", "تسعة", "عشرة", "أحد عشر", "اثنا عشر",
        "ثلاثة عشر", "أربعة عشر", "خمسة عشر", "ستة عشر",
        "سبعة عشر", "ثمانية عشر", "تسعة عشر",
    ]
    tens = ["", "", "عشرون", "ثلاثون", "أربعون", "خمسون",
            "ستون", "سبعون", "ثمانون", "تسعون"]
    hundreds = [
        "", "مائة", "مئتان", "ثلاثمائة", "أربعمائة", "خمسمائة",
        "ستمائة", "سبعمائة", "ثمانمائة", "تسعمائة",
    ]

    n = int(round(n))
    if n == 0:       return "صفر"
    if n < 0:        return "سالب " + tafqeet(-n)
    if n < 20:       return ones[n]
    if n < 100:
        t, o = divmod(n, 10)
        return tens[t] if o == 0 else ones[o] + " و" + tens[t]
    if n < 1_000:
        h, r = divmod(n, 100)
        return hundreds[h] if r == 0 else hundreds[h] + " و" + tafqeet(r)
    if n < 1_000_000:
        th, r = divmod(n, 1000)
        if   th == 1: pre = "ألف"
        elif th == 2: pre = "ألفان"
        else:         pre = tafqeet(th) + " آلاف"
        return pre if r == 0 else pre + " و" + tafqeet(r)
    return str(n)