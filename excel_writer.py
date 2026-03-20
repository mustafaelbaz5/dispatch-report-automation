"""
excel_writer.py — All Excel workbook / sheet writing logic.
No GUI or data processing here — pure formatting and output.
"""

from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font
from openpyxl.utils import get_column_letter

from config import (
    CENTER_TITLE, SECTOR_SHEETS, OUT_HEADERS, COL_WIDTHS,
    NAVY_FILL, GOLD_BG_FILL, ALT_FILL, WHITE_FILL, SUM_FILL, SIG_FILL,
    HDR_FONT, BIG_FONT, GOLD_FONT, SUM_FONT, DATA_FONT, SIG_FONT, DATE_FONT,
    BORDER, THICK_BORDER, BOT_THICK, TOP_THICK, THIN, THICK,
)
from processor import row_notes, tafqeet


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def _al(h="right", v="center", ro=2, wrap=False):
    """Shorthand for Alignment with RTL readingOrder."""
    return Alignment(horizontal=h, vertical=v, readingOrder=ro, wrap_text=wrap)


def _header_row(ws, row_num, headers, fill, font, border=None, height=24):
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=row_num, column=ci, value=h)
        c.font      = font
        c.fill      = fill
        c.alignment = _al("center")
        if border:
            c.border = border
    ws.row_dimensions[row_num].height = height


# ─────────────────────────────────────────────
# البيانات SHEET  (raw source data)
# ─────────────────────────────────────────────
def write_bianaat(wb: Workbook, raw_df: pd.DataFrame):
    ws = wb.active
    ws.title = "البيانات"
    ws.sheet_view.rightToLeft = True

    # Header row
    for ci, h in enumerate(raw_df.columns, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font      = HDR_FONT
        c.fill      = NAVY_FILL
        c.alignment = _al("center", wrap=True)
        c.border    = BORDER
    ws.row_dimensions[1].height = 22

    # Data rows
    for ri, row in raw_df.iterrows():
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri + 2, column=ci, value=val)
            c.font      = DATA_FONT
            c.border    = BORDER
            c.alignment = _al("right")
        ws.row_dimensions[ri + 2].height = 16

    # Auto column width (capped at 40)
    for ci in range(1, raw_df.shape[1] + 1):
        mx = max(
            (len(str(ws.cell(row=r, column=ci).value or ""))
             for r in range(1, raw_df.shape[0] + 2)),
            default=8,
        )
        ws.column_dimensions[get_column_letter(ci)].width = min(mx + 2, 40)

    ws.freeze_panes = "A2"


# ─────────────────────────────────────────────
# SECTOR SHEET  (بحرى / قبلى / etc.)
# ─────────────────────────────────────────────
def write_sector(wb: Workbook, sdf: pd.DataFrame, sector: str) -> tuple[int, float]:
    ws = wb.create_sheet(title=sector)
    ws.sheet_view.rightToLeft = True
    today_str = date.today().strftime("%Y-%m-%d")

    # ── Row 1: Institution title ──
    ws.merge_cells("A1:F1")
    c = ws.cell(row=1, column=1, value=CENTER_TITLE)
    c.font      = BIG_FONT
    c.fill      = NAVY_FILL
    c.alignment = _al("center")
    c.border    = THICK_BORDER
    ws.row_dimensions[1].height = 38

    # ── Row 2: Report + sector + date ──
    ws.merge_cells("A2:F2")
    c = ws.cell(row=2, column=1,
                value=f"بيان تسليم الارساليات الصادرة  ◈  {sector}  ◈  {today_str}")
    c.font      = GOLD_FONT
    c.fill      = GOLD_BG_FILL
    c.alignment = _al("center")
    c.border    = Border(left=THICK, right=THICK, top=THIN, bottom=THICK)
    ws.row_dimensions[2].height = 30

    # ── Row 3: Column headers ──
    _header_row(ws, 3, OUT_HEADERS, NAVY_FILL, HDR_FONT, border=BORDER, height=24)

    # ── Sort by dispatch number ──
    sdf = sdf.copy()
    sdf["_sort"] = pd.to_numeric(sdf["dispatch_no"], errors="coerce")
    sdf = sdf.sort_values("_sort").reset_index(drop=True)

    # ── Data rows ──
    for i, row in sdf.iterrows():
        er   = i + 4
        fill = ALT_FILL if i % 2 == 0 else WHITE_FILL
        vals = [
            i + 1,
            row["office_code"],
            row["office_name"],
            row["dispatch_no"],
            row["weight"],
            row_notes(row["total_items"]),
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.font      = DATA_FONT
            c.fill      = fill
            c.border    = BORDER
            c.alignment = _al("center" if ci in (1, 2, 4, 5) else "right")
        ws.row_dimensions[er].height = 18

    total_count  = len(sdf)
    total_weight = sdf["weight"].sum()
    sr           = total_count + 4

    # ── Summary: count ──
    ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=3)
    c1 = ws.cell(row=sr, column=1,
                 value=f"إجمالي عدد الارساليات :  {total_count}  ( {tafqeet(total_count)} )")
    c1.font      = SUM_FONT
    c1.fill      = SUM_FILL
    c1.border    = BOT_THICK
    c1.alignment = _al("right")

    # ── Summary: weight ──
    ws.merge_cells(start_row=sr, start_column=4, end_row=sr, end_column=6)
    c2 = ws.cell(row=sr, column=4,
                 value=f"إجمالي الوزن :  {total_weight:.3f} كجم  ( {tafqeet(total_weight)} كيلوجرام )")
    c2.font      = SUM_FONT
    c2.fill      = SUM_FILL
    c2.border    = BOT_THICK
    c2.alignment = _al("right")
    ws.row_dimensions[sr].height = 26

    # ── Blank spacer ──
    ws.row_dimensions[sr + 1].height = 10

    # ── Signature row ──
    sr3 = sr + 2
    ws.merge_cells(start_row=sr3, start_column=1, end_row=sr3, end_column=3)
    cs1 = ws.cell(row=sr3, column=1, value="توقيع المسلّم :  .................................")
    cs1.font      = SIG_FONT
    cs1.fill      = SIG_FILL
    cs1.border    = BORDER
    cs1.alignment = _al("right")

    ws.merge_cells(start_row=sr3, start_column=4, end_row=sr3, end_column=6)
    cs2 = ws.cell(row=sr3, column=4, value="توقيع المستلم :  .................................")
    cs2.font      = SIG_FONT
    cs2.fill      = SIG_FILL
    cs2.border    = BORDER
    cs2.alignment = _al("right")
    ws.row_dimensions[sr3].height = 30

    # ── Column widths & freeze ──
    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A4"

    return total_count, total_weight


# ─────────────────────────────────────────────
# الاجمالى SHEET  (grand summary)
# ─────────────────────────────────────────────
def write_ijmaly(wb: Workbook, sector_totals: dict):
    ws = wb.create_sheet(title="الاجمالى")
    ws.sheet_view.rightToLeft = True
    today_str = date.today().strftime("%Y-%m-%d")

    # ── Row 1: Institution title ──
    ws.merge_cells("B1:E1")
    c = ws.cell(row=1, column=2, value=CENTER_TITLE)
    c.font      = BIG_FONT
    c.fill      = NAVY_FILL
    c.border    = THICK_BORDER
    c.alignment = _al("center")
    ws.row_dimensions[1].height = 38

    # ── Row 2: Report title ──
    ws.merge_cells("B2:E2")
    c = ws.cell(row=2, column=2, value="بيان اجمالى الارساليات الصادرة")
    c.font      = GOLD_FONT
    c.fill      = GOLD_BG_FILL
    c.border    = Border(left=THICK, right=THICK, top=THIN, bottom=THICK)
    c.alignment = _al("center")
    ws.row_dimensions[2].height = 30

    # ── Row 3: Date ──
    ws.merge_cells("B3:E3")
    c = ws.cell(row=3, column=2, value=f"التاريخ :  {today_str}")
    c.font      = DATE_FONT
    c.fill      = WHITE_FILL
    c.alignment = _al("center")
    ws.row_dimensions[3].height = 22

    # ── Row 4: Spacer ──
    ws.row_dimensions[4].height = 8

    # ── Row 5: Headers ──
    ijmaly_headers = ["م", "القطاع", "عدد الارساليات", "الوزن (كجم)"]
    for ci, h in enumerate(ijmaly_headers, 2):
        c = ws.cell(row=5, column=ci, value=h)
        c.font      = HDR_FONT
        c.fill      = NAVY_FILL
        c.border    = BORDER
        c.alignment = _al("center")
    ws.row_dimensions[5].height = 24

    # ── Data rows ──
    grand_count = 0
    grand_weight = 0.0
    for i, sector in enumerate(SECTOR_SHEETS, 1):
        info  = sector_totals.get(sector, {"count": 0, "weight": 0.0})
        er    = i + 5
        fill  = ALT_FILL if i % 2 == 0 else WHITE_FILL
        grand_count  += info["count"]
        grand_weight += info["weight"]
        for ci, val in enumerate([i, sector, info["count"], round(info["weight"], 3)], 2):
            c = ws.cell(row=er, column=ci, value=val)
            c.font      = DATA_FONT
            c.fill      = fill
            c.border    = BORDER
            c.alignment = _al("center")
        ws.row_dimensions[er].height = 22

    # ── Grand total ──
    tr = len(SECTOR_SHEETS) + 6
    for ci, val in enumerate(["الاجمالى", grand_count, round(grand_weight, 3)], 3):
        c = ws.cell(row=tr, column=ci, value=val)
        c.font      = SUM_FONT
        c.fill      = SUM_FILL
        c.border    = BOT_THICK
        c.alignment = _al("center")
    ws.row_dimensions[tr].height = 26

    # ── Column widths ──
    for ci, w in zip([2, 3, 4, 5], [6, 20, 22, 18]):
        ws.column_dimensions[get_column_letter(ci)].width = w


# ─────────────────────────────────────────────
# MAIN EXPORT FUNCTION
# ─────────────────────────────────────────────
def build_workbook(raw_df: pd.DataFrame, filtered_df: pd.DataFrame,
                   output_dir: str, log_fn) -> str:
    """Build the full workbook and save it. Returns the output path."""
    log_fn("📝 إنشاء ملف Excel...")
    wb = Workbook()

    write_bianaat(wb, raw_df)

    sector_totals = {}
    for sector in SECTOR_SHEETS:
        sdf = filtered_df[filtered_df["sector"] == sector].copy()
        log_fn(f"   📊 {sector}: {len(sdf)} ارسالية")
        cnt, wgt = write_sector(wb, sdf, sector)
        sector_totals[sector] = {"count": cnt, "weight": wgt}

    write_ijmaly(wb, sector_totals)

    # ── Build output path with auto-increment ──
    save_dir = Path(output_dir)
    save_dir.mkdir(parents=True, exist_ok=True)
    today    = date.today().strftime("%Y-%m-%d")
    base     = f"بيان تسليم الارساليات الصادرة - {today}"
    out_path = save_dir / f"{base}.xlsx"
    if out_path.exists():
        i = 1
        while (save_dir / f"{base} ({i}).xlsx").exists():
            i += 1
        out_path = save_dir / f"{base} ({i}).xlsx"

    wb.save(out_path)
    log_fn(f"✅ تم الحفظ:\n{out_path}")
    return str(out_path)