"""
processor.py — Data loading, cleaning, filtering, and mapping.
No GUI or Excel writing here — pure data logic.
"""

import re
from pathlib import Path

import pandas as pd

from config import (
    COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
    COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT,
    VALID_STATUSES,
)


def extract_code(val) -> str:
    """Extract numeric office code from '3000-Alexandria...' → '3000'."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    m = re.match(r"^(\d+)", str(val).strip())
    return m.group(1) if m else str(val).strip()


def tafqeet(n: float) -> str:
    """Convert number to Arabic words (تفقيط)."""
    ones = ["","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة","عشرة",
            "أحد عشر","اثنا عشر","ثلاثة عشر","أربعة عشر","خمسة عشر","ستة عشر",
            "سبعة عشر","ثمانية عشر","تسعة عشر"]
    tens = ["","","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"]
    n = int(round(n))
    if n == 0: return "صفر"
    if n < 0:  return "سالب " + tafqeet(-n)
    if n < 20: return ones[n]
    if n < 100:
        t, o = divmod(n, 10)
        return tens[t] if o == 0 else ones[o] + " و" + tens[t]
    if n < 1000:
        hh = ["","مائة","مئتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"]
        h, r = divmod(n, 100)
        return hh[h] if r == 0 else hh[h] + " و" + tafqeet(r)
    if n < 1_000_000:
        th, r = divmod(n, 1000)
        pre = "ألف" if th == 1 else ("ألفان" if th == 2 else tafqeet(th) + " آلاف")
        return pre if r == 0 else pre + " و" + tafqeet(r)
    return str(n)


def row_notes(total_items: float) -> str:
    """Return shipment notes based on item count."""
    return "على المكشوف - قابل للكسر" if total_items == 1 else "قابل للكسر"


def load_mapping(script_dir: Path) -> tuple[dict, set]:
    """
    Load القطاعات.xlsx from the script folder.
    Returns: (mapping dict, set of valid codes)
    mapping = { "3000": {"office_name": "إسكندرية", "sector": "بحرى"}, ... }
    """
    path = Path(script_dir) / "القطاعات.xlsx"
    if not path.exists():
        raise RuntimeError(f"ملف القطاعات.xlsx غير موجود في:\n{script_dir}")
    df = pd.read_excel(path, header=0, dtype=str)
    mapping = {}
    for _, row in df.iterrows():
        code = str(row.iloc[0]).strip()
        if code and code != "nan":
            mapping[code] = {
                "office_name": str(row.iloc[1]).strip(),
                "sector":      str(row.iloc[2]).strip(),
            }
    return mapping, set(mapping.keys())


def load_and_filter(source_path: str, after_6pm: bool, log_fn) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load source Excel, apply all filters, attach sector/office_name.
    Returns: (raw_df for البيانات sheet, filtered_df for sector sheets)
    """
    log_fn("📂 قراءة ملف المصدر...")
    try:
        xl = pd.ExcelFile(source_path)
    except Exception as e:
        raise RuntimeError(f"تعذّر فتح الملف: {e}")

    # Raw data — used as-is for البيانات sheet
    raw_df = pd.read_excel(xl, sheet_name=0, header=0, dtype=str)

    # Working data — real column headers at row index 1
    main_df = pd.read_excel(xl, sheet_name=0, header=1, dtype=str)

    log_fn("📋 تحميل ملف القطاعات...")
    script_dir = Path(__file__).parent
    mapping, valid_codes = load_mapping(script_dir)
    log_fn(f"   ✓ {len(valid_codes)} كود محمّل")

    log_fn("🔧 تنظيف البيانات...")
    needed = max(COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
                 COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT)
    if main_df.shape[1] <= needed:
        raise RuntimeError(f"الملف يحتوي {main_df.shape[1]} عمود فقط، مطلوب {needed+1}.")

    df = main_df.iloc[:, [
        COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
        COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT
    ]].copy()
    df.columns = ["dest_office", "dispatch_no", "dt_raw", "status", "total_items", "weight"]

    df["office_code"] = df["dest_office"].apply(extract_code)
    df["datetime"]    = pd.to_datetime(df["dt_raw"], errors="coerce")
    df["total_items"] = pd.to_numeric(df["total_items"], errors="coerce").fillna(0)
    df["weight"]      = pd.to_numeric(df["weight"], errors="coerce").fillna(0)
    df["dispatch_no"] = df["dispatch_no"].astype(str).str.strip()
    df["status"]      = df["status"].astype(str).str.strip()

    log_fn(f"   إجمالي الصفوف: {len(df)}")

    # ── Apply filters ──
    df = df[df["status"].str.lower().isin(VALID_STATUSES)]
    log_fn(f"   ✓ {len(df)} بعد فلتر الحالة (Closed / FullyReceived)")

    df = df[df["total_items"] > 0]
    log_fn(f"   ✓ {len(df)} بعد فلتر العناصر > 0")

    df = df[df["office_code"].isin(valid_codes)]
    log_fn(f"   ✓ {len(df)} بعد فلتر الكود")

    if after_6pm:
        df = df[df["datetime"].apply(lambda d: pd.notna(d) and d.hour >= 18)]
        log_fn(f"   ✓ {len(df)} بعد فلتر ما بعد 6 مساءً")

    df = df.reset_index(drop=True)
    if df.empty:
        raise RuntimeError("لا توجد بيانات بعد تطبيق الفلاتر.")

    # Attach names from mapping
    df["sector"]      = df["office_code"].map(lambda c: mapping[c]["sector"])
    df["office_name"] = df["office_code"].map(lambda c: mapping[c]["office_name"])

    return raw_df, df