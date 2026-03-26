"""
config/report.py — Report titles, sector names, headers, and save path.
"""

from pathlib import Path

CENTER_TITLE  = "عمليات البريد السريع بالمنصورة 9001"

SECTOR_SHEETS = ["بحرى", "قبلى", "رمسيس", "طنطا", "المنصورة"]

OUT_HEADERS   = [
    "م", "الكود", "اسم مركز الحركة", "رقم الارسالية",
    "الوزن", "قابل للكسر", "على المكشوف",
]

NOTES_MERGE_HEADER = "ملاحظات"

DEFAULT_SAVE_DIR = Path.home() / "Desktop" / "التقارير"