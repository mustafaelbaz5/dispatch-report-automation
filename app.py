"""
app.py — Main GUI entry point.
Run: python app.py

Requirements:
    pip install customtkinter pandas openpyxl tkinterdnd2 pywin32
"""

import os
import threading
import subprocess
import sys
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

from config import DEFAULT_SAVE_DIR, SECTOR_SHEETS
from processor import load_and_filter
from excel_writer import build_workbook

# ─────────────────────────────────────────────
# PALETTE
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_APP     = "#0A0D13"
BG_SURFACE = "#111827"
BG_CARD    = "#141E2E"
BG_INPUT   = "#0F1825"
BG_HOVER   = "#1C2A3A"

BORDER     = "#1E3048"
BORDER_LT  = "#2A4060"

G          = "#10B981"
G_H        = "#059669"
G_DIM      = "#052E20"

T          = "#38BDF8"
T_H        = "#0EA5E9"
T_DIM      = "#0A2540"

A          = "#818CF8"
A_H        = "#6366F1"

TEXT_PRI   = "#F0F4F8"
TEXT_SEC   = "#64748B"
TEXT_DIM   = "#2D3F54"
TEXT_GOLD  = "#F59E0B"

RED        = "#F87171"
AMBER      = "#FBBF24"
SUCCESS    = "#34D399"


def _f(size=12, bold=False, family="Tajawal"):
    return ctk.CTkFont(family=family, size=size,
                       weight="bold" if bold else "normal")

def _mono(size=10):
    return ctk.CTkFont(family="Consolas", size=size)


# ─────────────────────────────────────────────
# REUSABLE WIDGETS
# ─────────────────────────────────────────────
class Divider(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BORDER, height=1, **kw)


class PrimaryBtn(ctk.CTkButton):
    def __init__(self, parent, text, cmd, icon="",
                 height=46, color=G, hcolor=None, **kw):
        lbl = f"{icon}  {text}" if icon else text
        super().__init__(parent, text=lbl, font=_f(13, bold=True),
                         fg_color=color, hover_color=hcolor or G_H,
                         text_color="#FFFFFF", height=height,
                         corner_radius=10, command=cmd, **kw)


class GhostBtn(ctk.CTkButton):
    def __init__(self, parent, text, cmd,
                 accent=T, icon="", height=38, **kw):
        lbl = f"{icon}  {text}" if icon else text
        super().__init__(parent, text=lbl, font=_f(11, bold=True),
                         fg_color="transparent", hover_color=BG_HOVER,
                         text_color=accent, height=height,
                         corner_radius=8, border_width=1,
                         border_color=accent, command=cmd, **kw)


# ─────────────────────────────────────────────
# PRINT DIALOG — sector sheets
# ─────────────────────────────────────────────
class PrintDialog(ctk.CTkToplevel):
    def __init__(self, parent, output_path: str):
        super().__init__(parent)
        self.title("طباعة الشيتات")
        self.geometry("420x480")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.lift(); self.focus_force(); self.grab_set()
        self._path   = output_path
        self._checks = {}
        self._build()

    def _build(self):
        ctk.CTkFrame(self, fg_color=G, height=3, corner_radius=0).pack(fill="x")

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16, 8))
        ctk.CTkLabel(hdr, text="🖨  اختر الشيتات للطباعة",
                     font=_f(15, bold=True), text_color=TEXT_PRI,
                     anchor="e").pack(anchor="e")
        ctk.CTkLabel(hdr, text=Path(self._path).name,
                     font=_f(9), text_color=TEXT_SEC,
                     anchor="e", wraplength=380).pack(anchor="e", pady=(2, 0))

        Divider(self).pack(fill="x", padx=20, pady=(8, 0))

        sr = ctk.CTkFrame(self, fg_color="transparent")
        sr.pack(fill="x", padx=20, pady=(8, 6))
        for lbl, state in [("تحديد الكل", True), ("إلغاء الكل", False)]:
            ctk.CTkButton(sr, text=lbl, font=_f(9),
                          fg_color=BG_INPUT, hover_color=BG_HOVER,
                          text_color=TEXT_SEC, border_width=1,
                          border_color=BORDER, height=26, width=84,
                          corner_radius=6,
                          command=lambda s=state: self._sel_all(s)
                          ).pack(side="left", padx=(0, 6))

        sf = ctk.CTkScrollableFrame(self, fg_color=BG_INPUT,
                                    corner_radius=10, border_width=1,
                                    border_color=BORDER, height=200)
        sf.pack(fill="x", padx=20, pady=(0, 12))

        for sheet in SECTOR_SHEETS + ["الاجمالى"]:
            var = ctk.BooleanVar(value=True)
            self._checks[sheet] = var
            row = ctk.CTkFrame(sf, fg_color="transparent")
            row.pack(fill="x", padx=6, pady=3)
            ctk.CTkCheckBox(row, text=sheet, variable=var,
                            font=_f(12), text_color=TEXT_PRI,
                            checkmark_color="#fff", fg_color=G,
                            hover_color=G_H, border_color=BORDER_LT,
                            corner_radius=5).pack(side="right", padx=4)

        Divider(self).pack(fill="x", padx=20, pady=(0, 12))

        br = ctk.CTkFrame(self, fg_color="transparent")
        br.pack(fill="x", padx=20, pady=(0, 18))
        GhostBtn(br, text="إلغاء", cmd=self.destroy,
                 accent=TEXT_SEC, height=42, icon="✕"
                 ).pack(side="left", padx=(0, 8))
        PrimaryBtn(br, text="طباعة", cmd=self._print,
                   icon="🖨", height=42
                   ).pack(side="right", fill="x", expand=True)

    def _sel_all(self, state):
        for v in self._checks.values(): v.set(state)

    def _print(self):
        selected = [s for s, v in self._checks.items() if v.get()]
        if not selected:
            messagebox.showwarning("تنبيه", "يرجى تحديد شيت واحد على الأقل.", parent=self)
            return
        if not Path(self._path).exists():
            messagebox.showerror("خطأ", "ملف الإخراج غير موجود.", parent=self)
            return
        try:
            self._do_print(selected)
            self.destroy()
            messagebox.showinfo("الطباعة", f"✅ تم إرسال {len(selected)} شيت للطابعة.")
        except Exception as e:
            messagebox.showerror("خطأ في الطباعة", str(e), parent=self)

    def _do_print(self, sheets: list):
        import openpyxl, tempfile
        wb = openpyxl.load_workbook(self._path)
        for name in list(wb.sheetnames):
            if name not in sheets:
                del wb[name]
        if not wb.sheetnames:
            raise RuntimeError("لم يتم العثور على الشيتات المحددة.")
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, prefix="print_")
        tmp.close()
        wb.save(tmp.name)
        if sys.platform == "win32":
            os.startfile(tmp.name)
        else:
            subprocess.run(["lp", tmp.name], check=True)


# ─────────────────────────────────────────────
# PRINT DIALOG — الاجمالى
# ─────────────────────────────────────────────
class IjmalyPrintDialog(ctk.CTkToplevel):
    def __init__(self, parent, output_path: str, sector_totals: dict):
        super().__init__(parent)
        self.title("طباعة الاجمالى")
        self.geometry("440x500")
        self.resizable(False, False)
        self.configure(fg_color=BG_CARD)
        self.lift(); self.focus_force(); self.grab_set()
        self._path          = output_path
        self._sector_totals = sector_totals
        self._checks        = {}
        self._build()

    def _build(self):
        ctk.CTkFrame(self, fg_color=A, height=3, corner_radius=0).pack(fill="x")

        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(16, 8))
        ctk.CTkLabel(hdr, text="📊  اختر قطاعات الاجمالى",
                     font=_f(15, bold=True), text_color=TEXT_PRI,
                     anchor="e").pack(anchor="e")
        ctk.CTkLabel(hdr, text="اختر القطاعات التي تريد إظهارها في جدول الاجمالى",
                     font=_f(9), text_color=TEXT_SEC, anchor="e").pack(anchor="e")

        Divider(self).pack(fill="x", padx=20, pady=(8, 0))

        sr = ctk.CTkFrame(self, fg_color="transparent")
        sr.pack(fill="x", padx=20, pady=(8, 6))
        for lbl, state in [("تحديد الكل", True), ("إلغاء الكل", False)]:
            ctk.CTkButton(sr, text=lbl, font=_f(9),
                          fg_color=BG_INPUT, hover_color=BG_HOVER,
                          text_color=TEXT_SEC, border_width=1,
                          border_color=BORDER, height=26, width=84,
                          corner_radius=6,
                          command=lambda s=state: self._sel_all(s)
                          ).pack(side="left", padx=(0, 6))

        sf = ctk.CTkScrollableFrame(self, fg_color=BG_INPUT,
                                    corner_radius=10, border_width=1,
                                    border_color=BORDER, height=230)
        sf.pack(fill="x", padx=20, pady=(0, 12))

        for sector in SECTOR_SHEETS:
            info = self._sector_totals.get(sector, {"count": 0, "weight": 0.0})
            var  = ctk.BooleanVar(value=True)
            self._checks[sector] = var
            row = ctk.CTkFrame(sf, fg_color="transparent")
            row.pack(fill="x", padx=6, pady=4)

            pill = ctk.CTkFrame(row, fg_color=T_DIM, corner_radius=6)
            pill.pack(side="left")
            ctk.CTkLabel(pill,
                         text=f"{info['count']} ·  {info['weight']:.1f} كجم",
                         font=_f(9), text_color=T).pack(padx=8, pady=4)

            ctk.CTkCheckBox(row, text=sector, variable=var,
                            font=_f(12, bold=True), text_color=TEXT_PRI,
                            checkmark_color="#fff", fg_color=A,
                            hover_color=A_H, border_color=BORDER_LT,
                            corner_radius=5).pack(side="right", padx=4)

        Divider(self).pack(fill="x", padx=20, pady=(0, 12))

        br = ctk.CTkFrame(self, fg_color="transparent")
        br.pack(fill="x", padx=20, pady=(0, 18))
        GhostBtn(br, text="إلغاء", cmd=self.destroy,
                 accent=TEXT_SEC, height=42, icon="✕"
                 ).pack(side="left", padx=(0, 8))
        PrimaryBtn(br, text="طباعة", cmd=self._print,
                   color=A, hcolor=A_H, icon="🖨", height=42
                   ).pack(side="right", fill="x", expand=True)

    def _sel_all(self, state):
        for v in self._checks.values(): v.set(state)

    def _print(self):
        selected = [s for s, v in self._checks.items() if v.get()]
        if not selected:
            messagebox.showwarning("تنبيه", "يرجى تحديد قطاع واحد على الأقل.", parent=self)
            return
        try:
            self._do_print(selected)
            self.destroy()
            messagebox.showinfo("الطباعة", "✅ تم إرسال الاجمالى للطابعة.")
        except Exception as e:
            messagebox.showerror("خطأ في الطباعة", str(e), parent=self)

    def _do_print(self, selected_sectors: list):
        from datetime import date
        import tempfile
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from config import (
            CENTER_TITLE, HDR_FONT, SUM_FONT, DATA_FONT,
            HEADER_FILL, ALT_FILL, ODD_FILL, WHITE_FILL,
            THICK_BORDER, BORDER as XLBORDER, C_BLACK,
        )

        def _al(h="center", v="center", ro=2, wrap=False):
            return Alignment(horizontal=h, vertical=v, readingOrder=ro, wrap_text=wrap)

        today_str = date.today().strftime("%Y-%m-%d")
        wb = Workbook(); ws = wb.active
        ws.title = "الاجمالى"; ws.sheet_view.rightToLeft = True

        ws.merge_cells("A1:D1")
        c = ws.cell(row=1, column=1, value=CENTER_TITLE)
        c.font = Font(bold=True, name="Arial", size=14, color=C_BLACK)
        c.fill = WHITE_FILL; c.alignment = _al("center")
        ws.row_dimensions[1].height = 24

        ws.merge_cells("A2:D2")
        c = ws.cell(row=2, column=1, value="بيان اجمالى الارساليات الصادرة")
        c.font = Font(bold=True, name="Arial", size=13, color=C_BLACK)
        c.fill = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
        c.alignment = _al("center"); ws.row_dimensions[2].height = 20

        ws.merge_cells("A3:D3")
        c = ws.cell(row=3, column=1, value=f"التاريخ :  {today_str}")
        c.font = Font(bold=True, name="Arial", size=12, color=C_BLACK)
        c.fill = WHITE_FILL; c.alignment = _al("center")
        ws.row_dimensions[3].height = 16

        for ci, h in enumerate(["الوزن (كجم)", "عدد الارساليات", "القطاع", "م"], 1):
            c = ws.cell(row=4, column=ci, value=h)
            c.font = HDR_FONT; c.fill = HEADER_FILL; c.alignment = _al("center")
        ws.row_dimensions[4].height = 18

        grand_count = 0; grand_weight = 0.0
        for i, sector in enumerate(selected_sectors, 1):
            info = self._sector_totals.get(sector, {"count": 0, "weight": 0.0})
            er   = i + 4
            fill = ALT_FILL if i % 2 == 0 else ODD_FILL
            grand_count  += info["count"]
            grand_weight += info["weight"]
            for ci, val in enumerate([round(info["weight"], 3), info["count"], sector, i], 1):
                c = ws.cell(row=er, column=ci, value=val)
                c.font = DATA_FONT; c.fill = fill
                c.border = XLBORDER; c.alignment = _al("center")
            ws.row_dimensions[er].height = 18

        tr = len(selected_sectors) + 5
        ws.merge_cells(start_row=tr, start_column=3, end_row=tr, end_column=4)
        for ci, val in enumerate([round(grand_weight, 3), grand_count, "الاجمالى"], 1):
            c = ws.cell(row=tr, column=ci, value=val)
            c.font = SUM_FONT; c.fill = HEADER_FILL
            c.border = THICK_BORDER; c.alignment = _al("center")
        ws.row_dimensions[tr].height = 22

        sig = tr + 1
        ws.merge_cells(start_row=sig, start_column=1, end_row=sig, end_column=4)
        sb = ws.cell(row=sig, column=1,
                     value="توقيع المستلم :  .................................")
        sb.font = Font(bold=True, name="Arial", size=14, color=C_BLACK)
        sb.fill = PatternFill("solid", start_color="F2F2F2", end_color="F2F2F2")
        sb.alignment = _al("right"); ws.row_dimensions[sig].height = 30

        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 7
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize   = 9
        ws.page_setup.fitToPage   = True
        ws.page_setup.fitToWidth  = 1
        ws.page_margins.left = ws.page_margins.right  = 0.30
        ws.page_margins.top  = ws.page_margins.bottom = 0.30
        ws.print_options.horizontalCentered = True
        ws.print_title_rows = "1:4"

        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, prefix="ijmaly_print_")
        tmp.close(); wb.save(tmp.name)
        if sys.platform == "win32":
            os.startfile(tmp.name)
        else:
            subprocess.run(["lp", tmp.name], check=True)


# ─────────────────────────────────────────────
# STATUS DOT
# ─────────────────────────────────────────────
class StatusDot(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._dot = ctk.CTkFrame(self, width=7, height=7,
                                 corner_radius=4, fg_color=TEXT_DIM)
        self._dot.pack(side="left", padx=(0, 5))
        self._lbl = ctk.CTkLabel(self, text="جاهز",
                                 font=_f(10), text_color=TEXT_DIM)
        self._lbl.pack(side="left")

    def set(self, text, color):
        self._dot.configure(fg_color=color)
        self._lbl.configure(text=text, text_color=color)


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        if _DND_AVAILABLE:
            try: TkinterDnD._require(self) # type: ignore
            except Exception: pass

        self.title("بيان تسليم الارساليات الصادرة")
        self.configure(fg_color=BG_APP)

        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{sw // 2}x{sh}+0+0")
        self.minsize(580, 600)
        self.resizable(True, True)

        self._source_path   = ctk.StringVar()
        self._output_dir    = ctk.StringVar(value=str(DEFAULT_SAVE_DIR))
        self._after_6pm     = ctk.BooleanVar(value=False)
        self._last_output   = None
        self._sector_totals = {}
        self._log_visible   = False

        self._build_ui()

    # ─────────────────────────────────────────
    # UI
    # ─────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color=BG_APP,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=TEXT_SEC)
        self._scroll.pack(fill="both", expand=True)
        self._build_dropzone_card()
        self._build_output_row()
        self._build_options_row()
        self._build_actions()
        self._build_progress_row()
        self._build_log_panel()

    # ─────────────────────────────────────────
    # HEADER
    # ─────────────────────────────────────────
    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_SURFACE,
                           corner_radius=0, height=82)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        strip = ctk.CTkFrame(hdr, fg_color="transparent",
                             height=3, corner_radius=0)
        strip.pack(fill="x", side="top")
        strip.pack_propagate(False)
        ctk.CTkFrame(strip, fg_color=G, height=3,
                     corner_radius=0).pack(side="left", fill="both", expand=True)
        ctk.CTkFrame(strip, fg_color=T, height=3,
                     width=56, corner_radius=0).pack(side="left")
        ctk.CTkFrame(strip, fg_color=A, height=3,
                     width=28, corner_radius=0).pack(side="left")

        inner = ctk.CTkFrame(hdr, fg_color="transparent")
        inner.place(relx=0.5, rely=0.58, anchor="center")

        ctk.CTkLabel(inner,
                     text="بيان تسليم الارساليات الصادرة",
                     font=_f(18, bold=True),
                     text_color=TEXT_PRI).pack()

        sub = ctk.CTkFrame(inner, fg_color="transparent")
        sub.pack(pady=(3, 0))

        badge = ctk.CTkFrame(sub, fg_color=G_DIM,
                             corner_radius=5, height=17, width=34)
        badge.pack(side="right", padx=(6, 0))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text="9900",
                     font=_f(8, bold=True),
                     text_color=SUCCESS).pack(expand=True)

        ctk.CTkLabel(sub,
                     text="تتم الإدارة بواسطة رئيس العمليات البريدية المتخصصة"
                          " والمشرف على المركز اللوجيستى :  محمد شعبان",
                     font=_f(9), text_color=TEXT_GOLD).pack(side="right")

    # ─────────────────────────────────────────
    # DRAG & DROP CARD
    # ─────────────────────────────────────────
    def _build_dropzone_card(self):
        card = ctk.CTkFrame(self._scroll, fg_color=BG_CARD,
                            corner_radius=12, border_width=1,
                            border_color=BORDER)
        card.pack(fill="x", padx=14, pady=(12, 0))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=12)

        ctk.CTkLabel(inner, text="ملف المصدر ( Excel )",
                     font=_f(11, bold=True), text_color=TEXT_SEC,
                     anchor="e").pack(anchor="e", pady=(0, 8))

        # Drop zone
        self._dz = ctk.CTkFrame(inner, fg_color=BG_INPUT,
                                corner_radius=10, border_width=2,
                                border_color=BORDER, height=130)
        self._dz.pack(fill="x", pady=(0, 10))
        self._dz.pack_propagate(False)

        dz_c = ctk.CTkFrame(self._dz, fg_color="transparent")
        dz_c.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(dz_c, text="⬆", font=_f(20),
                     text_color=TEXT_DIM).pack()
        self._dz_lbl = ctk.CTkLabel(dz_c,
                                    text="اسحب ملف Excel وأفلته هنا",
                                    font=_f(12, bold=True),
                                    text_color=TEXT_SEC)
        self._dz_lbl.pack(pady=(2, 0))
        self._dz_sub = ctk.CTkLabel(dz_c,
                                    text=".xlsx  ·  .xls  ·  .xlsm",
                                    font=_f(9), text_color=TEXT_DIM)
        self._dz_sub.pack()

        if _DND_AVAILABLE:
            try:
                dz = self._dz._canvas          # type: ignore
                dz.drop_target_register(DND_FILES) # type: ignore
                dz.dnd_bind("<<Drop>>", self._on_drop) # type: ignore
            except Exception:
                pass
        self._dz.bind("<Enter>", lambda e: self._dz_hover(True))
        self._dz.bind("<Leave>", lambda e: self._dz_hover(False))

        # Path entry + browse
        pr = ctk.CTkFrame(inner, fg_color="transparent")
        pr.pack(fill="x")
        ctk.CTkEntry(pr, textvariable=self._source_path,
                     placeholder_text="أو اكتب / الصق المسار هنا ...",
                     font=_f(10), fg_color=BG_INPUT,
                     border_color=BORDER, text_color=TEXT_PRI,
                     height=36, justify="right", corner_radius=8
                     ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        GhostBtn(pr, text="تصفح", cmd=self._browse_source,
                 accent=T, height=36, icon="📁").pack(side="left")

    # ─────────────────────────────────────────
    # OUTPUT DIR
    # ─────────────────────────────────────────
    def _build_output_row(self):
        card = ctk.CTkFrame(self._scroll, fg_color=BG_CARD,
                            corner_radius=10, border_width=1,
                            border_color=BORDER)
        card.pack(fill="x", padx=14, pady=(8, 0))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(inner, text="💾  مجلد الحفظ",
                     font=_f(11, bold=True), text_color=TEXT_SEC,
                     anchor="e").pack(anchor="e", pady=(0, 6))

        or_ = ctk.CTkFrame(inner, fg_color="transparent")
        or_.pack(fill="x")
        ctk.CTkEntry(or_, textvariable=self._output_dir,
                     font=_f(10), fg_color=BG_INPUT,
                     border_color=BORDER, text_color=TEXT_PRI,
                     height=36, justify="right", corner_radius=8
                     ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        GhostBtn(or_, text="تصفح", cmd=self._browse_output,
                 accent=TEXT_SEC, height=36, icon="📂").pack(side="left")

    # ─────────────────────────────────────────
    # OPTIONS
    # ─────────────────────────────────────────
    def _build_options_row(self):
        card = ctk.CTkFrame(self._scroll, fg_color=BG_CARD,
                            corner_radius=10, border_width=1,
                            border_color=BORDER)
        card.pack(fill="x", padx=14, pady=(8, 0))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(inner,
                     text="⏰  تصفية ارساليات ما بعد 6 مساءً فقط",
                     font=_f(11), text_color=TEXT_PRI,
                     anchor="e").pack(side="right")
        ctk.CTkSwitch(inner, text="", variable=self._after_6pm,
                      progress_color=G, button_color=TEXT_PRI,
                      button_hover_color=SUCCESS, width=46
                      ).pack(side="left")

    # ─────────────────────────────────────────
    # ACTIONS
    # ─────────────────────────────────────────
    def _build_actions(self):
        wrap = ctk.CTkFrame(self._scroll, fg_color="transparent")
        wrap.pack(fill="x", padx=14, pady=(10, 0))

        self._start_btn = PrimaryBtn(wrap, text="ابدأ المعالجة",
                                     cmd=self._run, icon="▶", height=48)
        self._start_btn.pack(fill="x", pady=(0, 8))

        row = ctk.CTkFrame(wrap, fg_color="transparent")
        row.pack(fill="x", pady=(0, 6))
        row.columnconfigure(0, weight=1)
        row.columnconfigure(1, weight=1)
        row.columnconfigure(2, weight=1)

        self._open_btn = GhostBtn(row, text="فتح الملف",
                                   cmd=self._open_output,
                                   accent=T, height=38, icon="📂")
        self._open_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self._open_btn.configure(state="disabled")

        self._print_btn = GhostBtn(row, text="طباعة الشيتات",
                                    cmd=self._open_print_dialog,
                                    accent=A, height=38, icon="🖨")
        self._print_btn.grid(row=0, column=1, sticky="ew", padx=4)
        self._print_btn.configure(state="disabled")

        self._print_ijmaly_btn = GhostBtn(row, text="طباعة الاجمالى",
                                           cmd=self._open_ijmaly_print_dialog,
                                           accent=TEXT_GOLD, height=38, icon="📊")
        self._print_ijmaly_btn.grid(row=0, column=2, sticky="ew", padx=(4, 0))
        self._print_ijmaly_btn.configure(state="disabled")

        GhostBtn(wrap, text="تصفير", cmd=self._reset,
                 accent=TEXT_SEC, height=34, icon="↺").pack(fill="x")

    # ─────────────────────────────────────────
    # PROGRESS + LOG TOGGLE
    # ─────────────────────────────────────────
    def _build_progress_row(self):
        pf = ctk.CTkFrame(self._scroll, fg_color="transparent")
        pf.pack(fill="x", padx=14, pady=(10, 0))

        self._progress = ctk.CTkProgressBar(
            pf, height=4, progress_color=G,
            fg_color=BORDER, corner_radius=2)
        self._progress.pack(fill="x", pady=(0, 5))
        self._progress.set(0)

        row = ctk.CTkFrame(pf, fg_color="transparent")
        row.pack(fill="x")

        self._status = StatusDot(row)
        self._status.pack(side="right")

        self._log_toggle_btn = ctk.CTkButton(
            row, text="▾  السجل", font=_f(9),
            fg_color="transparent", hover_color=BG_HOVER,
            text_color=TEXT_SEC, height=22, width=72,
            border_width=1, border_color=BORDER, corner_radius=6,
            command=self._toggle_log)
        self._log_toggle_btn.pack(side="left")

    # ─────────────────────────────────────────
    # LOG PANEL — collapsible, hidden by default
    # ─────────────────────────────────────────
    def _build_log_panel(self):
        self._log_panel = ctk.CTkFrame(
            self._scroll, fg_color=BG_SURFACE,
            corner_radius=10, border_width=1, border_color=BORDER)
        # NOT packed — shown on toggle

        lhdr = ctk.CTkFrame(self._log_panel, fg_color="transparent")
        lhdr.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(lhdr, text="◉  سجل العمليات",
                     font=_f(9, bold=True), text_color=TEXT_SEC
                     ).pack(side="right")
        ctk.CTkButton(lhdr, text="مسح", font=_f(8),
                      width=44, height=20, fg_color=BG_INPUT,
                      hover_color=BG_HOVER, text_color=TEXT_DIM,
                      border_width=1, border_color=BORDER, corner_radius=5,
                      command=self._clear_log).pack(side="left")

        self._log_box = ctk.CTkTextbox(
            self._log_panel, font=_mono(10),
            fg_color="transparent", text_color="#94A3B8",
            border_width=0, height=120, wrap="word")
        self._log_box.pack(fill="x", padx=10, pady=(0, 8))

        tags = {
            "start": ("#F59E0B", ("Consolas", 10, "bold")),
            "ok":    (SUCCESS,   None),
            "info":  (T,         None),
            "step":  ("#475569", None),
            "error": (RED,       ("Consolas", 10, "bold")),
            "path":  ("#7DD3FC", None),
        }
        for name, (fg, fnt) in tags.items():
            kw: dict = {"foreground": fg}
            if fnt: kw["font"] = fnt
            self._log_box._textbox.tag_config(name, **kw)
        self._log_box.configure(state="disabled")

    def _toggle_log(self):
        if self._log_visible:
            self._log_panel.pack_forget()
            self._log_visible = False
            self._log_toggle_btn.configure(text="▾  السجل")
        else:
            self._log_panel.pack(fill="x", padx=14, pady=(8, 10))
            self._log_visible = True
            self._log_toggle_btn.configure(text="▴  السجل")

    # ─────────────────────────────────────────
    # DROP ZONE HELPERS
    # ─────────────────────────────────────────
    def _dz_hover(self, entering: bool):
        if entering:
            self._dz.configure(border_color=T, fg_color="#0D1E35")
        else:
            if not self._source_path.get():
                self._dz.configure(border_color=BORDER, fg_color=BG_INPUT)

    def _dz_set_file(self, path: str):
        self._source_path.set(path)
        self._dz_lbl.configure(text=f"✓  {Path(path).name}", text_color=SUCCESS)
        self._dz_sub.configure(text=path, text_color=TEXT_DIM)
        self._dz.configure(border_color=G, fg_color="#091A10")

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
            self._dz_set_file(path)
        else:
            messagebox.showwarning("تنبيه",
                "يرجى إسقاط ملف Excel فقط\n(.xlsx / .xls / .xlsm)")

    def _browse_source(self):
        p = filedialog.askopenfilename(
            title="اختر ملف Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if p: self._dz_set_file(p)

    def _browse_output(self):
        p = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if p: self._output_dir.set(p)

    # ─────────────────────────────────────────
    # LOG
    # ─────────────────────────────────────────
    def _log(self, msg: str):
        m = msg.strip()
        if m.startswith("🚀"):           tag = "start"
        elif "✅" in m or "بنجاح" in m:  tag = "ok"
        elif "❌" in m or "خطأ" in m:    tag = "error"
        elif "📁" in m:                   tag = "path"
        elif m.startswith("   ✓"):       tag = "ok"
        elif m.startswith("   "):        tag = "step"
        else:                             tag = "info"

        self._log_box._textbox.configure(state="normal")
        self._log_box._textbox.insert("end", msg + "\n", tag)
        self._log_box._textbox.see("end")
        self._log_box._textbox.configure(state="disabled")
        self._log_box.configure(state="disabled")

        # Auto-show log when processing starts
        if not self._log_visible:
            self._toggle_log()

    def _clear_log(self):
        self._log_box._textbox.configure(state="normal")
        self._log_box._textbox.delete("1.0", "end")
        self._log_box._textbox.configure(state="disabled")

    # ─────────────────────────────────────────
    # RESET
    # ─────────────────────────────────────────
    def _reset(self):
        self._source_path.set("")
        self._dz_lbl.configure(text="اسحب ملف Excel وأفلته هنا",
                                text_color=TEXT_SEC)
        self._dz_sub.configure(text=".xlsx  ·  .xls  ·  .xlsm",
                                text_color=TEXT_DIM)
        self._dz.configure(border_color=BORDER, fg_color=BG_INPUT)
        self._last_output   = None
        self._sector_totals = {}
        self._open_btn.configure(state="disabled")
        self._print_btn.configure(state="disabled")
        self._print_ijmaly_btn.configure(state="disabled")
        self._progress.set(0)
        self._status.set("جاهز", TEXT_DIM)
        self._clear_log()

    # ─────────────────────────────────────────
    # DIALOG LAUNCHERS
    # ─────────────────────────────────────────
    def _open_print_dialog(self):
        if not self._last_output or not Path(self._last_output).exists():
            messagebox.showwarning("تنبيه", "يرجى تشغيل المعالجة أولاً.")
            return
        PrintDialog(self, self._last_output)

    def _open_ijmaly_print_dialog(self):
        if not self._last_output or not Path(self._last_output).exists():
            messagebox.showwarning("تنبيه", "يرجى تشغيل المعالجة أولاً.")
            return
        IjmalyPrintDialog(self, self._last_output, self._sector_totals)

    # ─────────────────────────────────────────
    # RUN
    # ─────────────────────────────────────────
    def _run(self):
        path    = self._source_path.get().strip()
        out_dir = self._output_dir.get().strip() or str(DEFAULT_SAVE_DIR)

        if not path:
            messagebox.showwarning("تنبيه", "يرجى اختيار ملف Excel أولاً.")
            return
        if not os.path.exists(path):
            messagebox.showerror("خطأ", f"الملف غير موجود:\n{path}")
            return

        Path(out_dir).mkdir(parents=True, exist_ok=True)
        self._start_btn.configure(state="disabled", text="⏳  جارٍ المعالجة...")
        self._open_btn.configure(state="disabled")
        self._print_btn.configure(state="disabled")
        self._print_ijmaly_btn.configure(state="disabled")
        self._progress.set(0)
        self._status.set("جارٍ المعالجة ...", AMBER)
        self._clear_log()
        self._log("🚀 بدء المعالجة...")

        def worker():
            try:
                self.after(0, self._progress.set, 0.15)
                raw_df, filtered_df = load_and_filter(
                    path, self._after_6pm.get(),
                    lambda m: self.after(0, self._log, m))
                self.after(0, self._progress.set, 0.60)
                out, sector_totals = build_workbook(
                    raw_df, filtered_df, out_dir,
                    lambda m: self.after(0, self._log, m),
                    self._after_6pm.get())
                self.after(0, self._progress.set, 1.0)
                self.after(0, self._on_success, out, sector_totals)
            except Exception as e:
                self.after(0, self._on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    # ─────────────────────────────────────────
    # CALLBACKS
    # ─────────────────────────────────────────
    def _on_success(self, out_path: str, sector_totals: dict):
        self._last_output   = out_path
        self._sector_totals = sector_totals
        self._start_btn.configure(state="normal", text="▶  ابدأ المعالجة")
        self._open_btn.configure(state="normal")
        self._print_btn.configure(state="normal")
        self._print_ijmaly_btn.configure(state="normal")
        self._status.set("اكتملت المعالجة بنجاح ✓", SUCCESS)
        self._log(f"📁 {out_path}")
        messagebox.showinfo("تم بنجاح ✅",
                            f"تم إنشاء التقرير بنجاح!\n\n{out_path}")

    def _on_error(self, msg: str):
        self._start_btn.configure(state="normal", text="▶  ابدأ المعالجة")
        self._progress.set(0)
        self._status.set("حدث خطأ", RED)
        self._log(f"❌ خطأ: {msg}")
        messagebox.showerror("خطأ", msg)

    def _open_output(self):
        if self._last_output and os.path.exists(self._last_output):
            if sys.platform == "win32":
                os.startfile(self._last_output)
            elif sys.platform == "darwin":
                subprocess.call(["open", self._last_output])
            else:
                subprocess.call(["xdg-open", self._last_output])


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()