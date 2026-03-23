"""
core/processor.py — Data loading, cleaning, filtering, and sector mapping.
No GUI or Excel writing — pure data logic.
"""

import sys
from pathlib import Path

import pandas as pd

from config.columns import (
    COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
    COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT,
    VALID_STATUSES,
)
from utils.arabic import extract_code


def _resource_path(filename: str) -> Path:
    """Resolve path whether running as script or PyInstaller exe."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent.parent))
    return base / filename


def load_mapping() -> tuple[dict, set]:
    """
    Load القطاعات.xlsx from the project root.

    Returns:
        mapping  — { "3000": {"office_name": "...", "sector": "..."}, ... }
        valid_codes — set of all known codes
    """
    path = _resource_path("القطاعات.xlsx")
    if not path.exists():
        raise RuntimeError(f"ملف القطاعات.xlsx غير موجود في:\n{path.parent}")

    df = pd.read_excel(path, header=0, dtype=str)
    mapping = {
        str(row.iloc[0]).strip(): {
            "office_name": str(row.iloc[1]).strip(),
            "sector":      str(row.iloc[2]).strip(),
        }
        for _, row in df.iterrows()
        if str(row.iloc[0]).strip() not in ("", "nan")
    }
    return mapping, set(mapping.keys())


def _select_columns(main_df: pd.DataFrame) -> pd.DataFrame:
    """Pick the six needed columns and rename them."""
    needed = max(COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
                 COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT)
    if main_df.shape[1] <= needed:
        raise RuntimeError(
            f"الملف يحتوي {main_df.shape[1]} عمود فقط، مطلوب {needed + 1}."
        )
    df = main_df.iloc[:, [
        COL_DEST_OFFICE, COL_DISPATCH_NO, COL_DATETIME,
        COL_STATUS, COL_TOTAL_ITEMS, COL_WEIGHT,
    ]].copy()
    df.columns = ["dest_office", "dispatch_no", "dt_raw",
                  "status", "total_items", "weight"]
    return df


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """Parse and coerce all columns to their proper types."""
    df["office_code"] = df["dest_office"].apply(extract_code)
    df["datetime"]    = pd.to_datetime(df["dt_raw"], errors="coerce")
    df["total_items"] = pd.to_numeric(df["total_items"], errors="coerce").fillna(0)
    df["weight"]      = pd.to_numeric(df["weight"],      errors="coerce").fillna(0)
    df["dispatch_no"] = df["dispatch_no"].astype(str).str.strip()
    df["status"]      = df["status"].astype(str).str.strip()
    return df


def _apply_filters(df: pd.DataFrame, after_6pm: bool, log_fn) -> pd.DataFrame:
    """Apply all business filters and log each step."""
    log_fn(f"   إجمالي الصفوف: {len(df)}")

    df = df[df["status"].str.lower().isin(VALID_STATUSES)]
    log_fn(f"   ✓ {len(df)} بعد فلتر الحالة (Closed فقط)")

    df = df[df["total_items"] > 0]
    log_fn(f"   ✓ {len(df)} بعد فلتر العناصر > 0")

    return df


def _filter_by_codes(df: pd.DataFrame, valid_codes: set, log_fn) -> pd.DataFrame:
    df = df[df["office_code"].isin(valid_codes)]
    log_fn(f"   ✓ {len(df)} بعد فلتر الكود")
    return df


def _filter_after_6pm(df: pd.DataFrame, log_fn) -> pd.DataFrame:
    df = df[df["datetime"].apply(lambda d: pd.notna(d) and d.hour >= 18)]
    log_fn(f"   ✓ {len(df)} بعد فلتر ما بعد 6 مساءً")
    return df


def load_and_filter(
    source_path: str,
    after_6pm: bool,
    log_fn,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load source Excel, apply all filters, attach sector / office_name.

    Returns:
        raw_df      — unmodified first sheet (used for البيانات sheet)
        filtered_df — cleaned, filtered, sector-mapped rows
    """
    log_fn("📂 قراءة ملف المصدر...")
    try:
        xl = pd.ExcelFile(source_path)
    except Exception as e:
        raise RuntimeError(f"تعذّر فتح الملف: {e}")

    raw_df  = pd.read_excel(xl, sheet_name=0, header=0, dtype=str)
    main_df = pd.read_excel(xl, sheet_name=0, header=1, dtype=str)

    log_fn("📋 تحميل ملف القطاعات...")
    mapping, valid_codes = load_mapping()
    log_fn(f"   ✓ {len(valid_codes)} كود محمّل")

    log_fn("🔧 تنظيف البيانات...")
    df = _clean(_select_columns(main_df))
    df = _apply_filters(df, after_6pm, log_fn)
    df = _filter_by_codes(df, valid_codes, log_fn)

    if after_6pm:
        df = _filter_after_6pm(df, log_fn)

    df = df.reset_index(drop=True)
    if df.empty:
        raise RuntimeError("لا توجد بيانات بعد تطبيق الفلاتر.")

    df["sector"]      = df["office_code"].map(lambda c: mapping[c]["sector"])
    df["office_name"] = df["office_code"].map(lambda c: mapping[c]["office_name"])

    return raw_df, df