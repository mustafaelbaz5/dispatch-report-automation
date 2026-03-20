"""
app.py — Main GUI entry point.
Run this file to launch the application.

Requirements:
    pip install customtkinter pandas openpyxl tkinterdnd2

Folder structure:
    📁 shipment_app/
        app.py            ← run this
        config.py
        processor.py
        excel_writer.py
        القطاعات.xlsx     ← mapping file must be here
"""

import os
import threading
import subprocess
import sys
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

# Optional drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

from config import DEFAULT_SAVE_DIR
from processor import load_and_filter
from excel_writer import build_workbook


# ─────────────────────────────────────────────
# THEME & PALETTE
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_APP      = "#0D1117"
BG_CARD     = "#161B22"
BG_INPUT    = "#21262D"
BG_HEADER   = "#0D1F3C"
ACCENT      = "#388BFD"
ACCENT_H    = "#1F6FEB"
ACCENT_GLOW = "#1C3A6E"
GREEN       = "#3FB950"
GREEN_H     = "#2EA043"
GREEN_DIM   = "#1B4332"
AMBER       = "#D29922"
RED         = "#F85149"
BORDER_C    = "#30363D"
TEXT_PRI    = "#E6EDF3"
TEXT_SEC    = "#8B949E"
TEXT_DIM    = "#484F58"
TEXT_GOLD   = "#E3B341"
STAT_COLORS = ["#388BFD", "#3FB950", "#E3B341", "#F78166", "#BC8CFF"]


def _font(size=12, bold=False):
    return ctk.CTkFont(family="Tajawal", size=size,
                       weight="bold" if bold else "normal")


# ─────────────────────────────────────────────
# WIDGETS
# ─────────────────────────────────────────────
class SectionLabel(ctk.CTkLabel):
    def __init__(self, parent, text, size=11, color=TEXT_SEC, **kw):
        super().__init__(parent, text=text, font=_font(size),
                         text_color=color, anchor="e", justify="right", **kw)


class Divider(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BORDER_C, height=1, **kw)


class InputRow(ctk.CTkFrame):
    def __init__(self, parent, label, var, placeholder,
                 btn_text=None, btn_cmd=None, btn_color=ACCENT, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        SectionLabel(self, label, size=11).pack(anchor="e", pady=(0, 5))
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x")
        self.entry = ctk.CTkEntry(
            row, textvariable=var,
            placeholder_text=placeholder,
            font=_font(11), fg_color=BG_INPUT, border_color=BORDER_C,
            text_color=TEXT_PRI, height=40, justify="right", corner_radius=8,
        )
        self.entry.pack(side="left", fill="x", expand=True,
                        padx=(0, 8) if btn_text else 0)
        if btn_text:
            ctk.CTkButton(
                row, text=btn_text, font=_font(11, bold=True),
                fg_color=btn_color, hover_color=ACCENT_H,
                width=105, height=40, corner_radius=8,
                command=btn_cmd,
            ).pack(side="left")


class StatCard(ctk.CTkFrame):
    def __init__(self, parent, sector, count, color=ACCENT, **kw):
        super().__init__(parent, fg_color=BG_INPUT, corner_radius=10,
                         border_width=1, border_color=color, **kw)
        ctk.CTkFrame(self, fg_color=color, height=4, corner_radius=0).pack(fill="x")
        ctk.CTkLabel(self, text=sector, font=_font(12, bold=True),
                     text_color=TEXT_PRI).pack(pady=(10, 2))
        ctk.CTkLabel(self, text=str(count), font=_font(22, bold=True),
                     text_color=color).pack()
        ctk.CTkLabel(self, text="ارسالية", font=_font(9),
                     text_color=TEXT_DIM).pack(pady=(0, 10))


class CollapsibleSection(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._visible = False
        self.pack_forget()

    def show(self):
        if not self._visible:
            self._visible = True
            self.pack(fill="x", padx=24, pady=(0, 12))

    def hide(self):
        if self._visible:
            self._visible = False
            self.pack_forget()


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        if _DND_AVAILABLE:
            try:
                TkinterDnD._require(self)  # type: ignore
            except Exception:
                pass

        self.title("بيان تسليم الارساليات الصادرة")
        self.geometry("800x720")
        self.minsize(700, 600)
        self.configure(fg_color=BG_APP)

        self._source_path = ctk.StringVar()
        self._output_dir  = ctk.StringVar(value=str(DEFAULT_SAVE_DIR))
        self._after_6pm   = ctk.BooleanVar(value=False)
        self._last_output = None

        self._build_ui()

    def _build_ui(self):
        self._build_header()
        self._build_body()

    # ─────────────────────────────────────────
    # HEADER
    # ─────────────────────────────────────────
    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_HEADER, corner_radius=0, height=96)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkFrame(hdr, fg_color=ACCENT, width=4, corner_radius=0).pack(
            side="left", fill="y")

        inner = ctk.CTkFrame(hdr, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(inner, text="بيان تسليم الارساليات الصادرة",
                     font=_font(22, bold=True), text_color=TEXT_PRI).pack()
        ctk.CTkLabel(inner,
                     text="المركز اللوجيستى بالمنصورة 9900",
                     font=_font(12), text_color=TEXT_GOLD).pack(pady=(3, 0))

    # ─────────────────────────────────────────
    # BODY
    # ─────────────────────────────────────────
    def _build_body(self):
        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color=BG_APP,
            scrollbar_button_color=BORDER_C,
            scrollbar_button_hover_color=TEXT_SEC,
        )
        self._scroll.pack(fill="both", expand=True)
        self._build_files_card()
        self._build_options_card()
        self._build_actions()
        self._build_log_section()
        self._build_stats_section()

    def _build_files_card(self):
        inner = self._card(self._scroll, "الملفات", "📁")

        SectionLabel(inner, "ملف المصدر  ( Excel )  :", size=11).pack(
            anchor="e", pady=(0, 6))

        # ── Drop zone ──
        self._drop_zone = ctk.CTkFrame(
            inner, fg_color=BG_INPUT, corner_radius=12,
            border_width=2, border_color=BORDER_C, height=180,
        )
        self._drop_zone.pack(fill="x", pady=(0, 8))
        self._drop_zone.pack_propagate(False)

        ctk.CTkLabel(self._drop_zone, text="☁", font=_font(18),
                     text_color=TEXT_DIM).place(relx=0.5, rely=0.30, anchor="center")

        self._drop_lbl = ctk.CTkLabel(
            self._drop_zone, text="اسحب ملف Excel وأفلته هنا",
            font=_font(10, bold=True), text_color=TEXT_SEC,
        )
        self._drop_lbl.place(relx=0.5, rely=0.57, anchor="center")

        self._drop_sub = ctk.CTkLabel(
            self._drop_zone, text="يدعم  .xlsx   .xls   .xlsm",
            font=_font(8), text_color=TEXT_PRI,
        )
        self._drop_sub.place(relx=0.5, rely=0.78, anchor="center")

        if _DND_AVAILABLE:
            try:
                dz = self._drop_zone._canvas  # type: ignore
                dz.drop_target_register(DND_FILES)  # type: ignore
                dz.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore
            except Exception:
                pass

        self._drop_zone.bind("<Enter>", lambda e: self._dz_hover(True))
        self._drop_zone.bind("<Leave>", lambda e: self._dz_hover(False))

        # ── Path entry + browse ──
        browse_row = ctk.CTkFrame(inner, fg_color="transparent")
        browse_row.pack(fill="x", pady=(0, 14))

        ctk.CTkEntry(
            browse_row, textvariable=self._source_path,
            placeholder_text="أو اكتب / الصق المسار هنا ...",
            font=_font(11), fg_color=BG_INPUT, border_color=BORDER_C,
            text_color=TEXT_PRI, height=40, justify="right", corner_radius=8,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            browse_row, text="📁  تصفح",
            font=_font(11, bold=True), fg_color=ACCENT, hover_color=ACCENT_H,
            width=105, height=40, corner_radius=8,
            command=self._browse_source,
        ).pack(side="left")

        Divider(inner).pack(fill="x", pady=(0, 14))

        InputRow(
            inner, "مجلد حفظ التقارير  :",
            self._output_dir, "مسار مجلد الحفظ ...",
            btn_text="📂  تصفح", btn_cmd=self._browse_output,
            btn_color="#3D4F6B",
        ).pack(fill="x")

    def _build_options_card(self):
        inner = self._card(self._scroll, "خيارات المعالجة", "⚙️")

        row = ctk.CTkFrame(inner, fg_color=BG_INPUT, corner_radius=8)
        row.pack(fill="x")

        ctk.CTkLabel(
            row,
            text="تصفية الارساليات المرسلة بعد الساعة 6 مساءً فقط",
            font=_font(12), text_color=TEXT_PRI,
        ).pack(side="right", padx=(0, 14), pady=13)

        ctk.CTkSwitch(
            row, text="", variable=self._after_6pm,
            progress_color=ACCENT, button_color=TEXT_PRI, width=46,
        ).pack(side="left", padx=14, pady=13)

    def _build_actions(self):
        wrapper = ctk.CTkFrame(self._scroll, fg_color="transparent")
        wrapper.pack(fill="x", padx=24, pady=(16, 6))

        # ── Two buttons side by side ──
        btn_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 10))
        btn_row.columnconfigure(0, weight=3)
        btn_row.columnconfigure(1, weight=2)

        self._start_btn = ctk.CTkButton(
            btn_row,
            text="▶   ابدأ المعالجة",
            font=_font(15, bold=True),
            fg_color=GREEN, hover_color=GREEN_H,
            height=52, corner_radius=12,
            command=self._run,
        )
        self._start_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._open_btn = ctk.CTkButton(
            btn_row,
            text="📂   فتح الملف",
            font=_font(13, bold=True),
            fg_color=ACCENT_GLOW, hover_color="#243A5E",
            text_color=ACCENT, height=52, corner_radius=12,
            border_width=1, border_color=ACCENT,
            command=self._open_output, state="disabled",
        )
        self._open_btn.grid(row=0, column=1, sticky="ew")

        # ── Progress bar ──
        prog_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        prog_row.pack(fill="x")

        self._progress = ctk.CTkProgressBar(
            prog_row, height=5, progress_color=ACCENT,
            fg_color=BORDER_C, corner_radius=3,
        )
        self._progress.pack(fill="x", pady=(0, 4))
        self._progress.set(0)

        self._status_lbl = ctk.CTkLabel(
            prog_row, text="● جاهز",
            font=_font(10), text_color=TEXT_DIM, anchor="e",
        )
        self._status_lbl.pack(anchor="e")

    def _build_log_section(self):
        self._log_section = CollapsibleSection(self._scroll)

        # ── Header row ──
        hdr = ctk.CTkFrame(self._log_section, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(hdr, text="سجل العمليات",
                     font=_font(13, bold=True),
                     text_color=TEXT_SEC).pack(side="right")
        ctk.CTkButton(
            hdr, text="🗑  مسح",
            font=_font(10), width=75, height=28,
            fg_color=BORDER_C, hover_color="#3D444D",
            text_color=TEXT_SEC, corner_radius=6,
            command=self._clear_log,
        ).pack(side="left")

        # ── Card ──
        card = ctk.CTkFrame(self._log_section, fg_color="#0D1117",
                            corner_radius=12,
                            border_width=1, border_color=BORDER_C)
        card.pack(fill="x")

        # Top colored bar
        ctk.CTkFrame(card, fg_color=ACCENT, height=3,
                     corner_radius=0).pack(fill="x")

        self._log_box = ctk.CTkTextbox(
            card,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color="#0D1117",
            text_color="#C9D1D9",   # default — overridden by tags
            border_width=0,
            height=180, wrap="word",
        )
        self._log_box.pack(fill="x", padx=12, pady=(8, 12))

        # ── Color tags for different message types ──
        self._log_box._textbox.tag_config(
            "start",  foreground="#E3B341", font=("Consolas", 12, "bold"))
        self._log_box._textbox.tag_config(
            "ok",     foreground="#3FB950")
        self._log_box._textbox.tag_config(
            "info",   foreground="#79C0FF")
        self._log_box._textbox.tag_config(
            "step",   foreground="#8B949E")
        self._log_box._textbox.tag_config(
            "save",   foreground="#58A6FF", font=("Consolas", 12, "bold"))
        self._log_box._textbox.tag_config(
            "error",  foreground="#F85149", font=("Consolas", 12, "bold"))
        self._log_box._textbox.tag_config(
            "path",   foreground="#D2A8FF")

        self._log_box.configure(state="disabled")

    def _build_stats_section(self):
        self._stats_section = CollapsibleSection(self._scroll)

        hdr = ctk.CTkFrame(self._stats_section, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(hdr, text="ملخص النتائج", font=_font(13, bold=True),
                     text_color=TEXT_SEC).pack(side="right")

        stats_card = ctk.CTkFrame(self._stats_section, fg_color=BG_CARD,
                                   corner_radius=12)
        stats_card.pack(fill="x")

        self._stats_inner = ctk.CTkFrame(stats_card, fg_color="transparent")
        self._stats_inner.pack(fill="x", padx=14, pady=14)

        self._total_lbl = ctk.CTkLabel(
            stats_card, text="", font=_font(13, bold=True), text_color=ACCENT,
        )
        self._total_lbl.pack(pady=(0, 14))

    # ─────────────────────────────────────────
    # CARD FACTORY
    # ─────────────────────────────────────────
    def _card(self, parent, title: str, icon: str = "") -> ctk.CTkFrame:
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="x", padx=24, pady=(16, 0))

        title_row = ctk.CTkFrame(wrapper, fg_color="transparent")
        title_row.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            title_row,
            text=f"{icon}  {title}" if icon else title,
            font=_font(14, bold=True), text_color=TEXT_PRI,
        ).pack(side="right")

        ctk.CTkFrame(title_row, fg_color=ACCENT, height=2,
                     corner_radius=1).pack(side="right", padx=(0, 8), pady=7)

        card = ctk.CTkFrame(wrapper, fg_color=BG_CARD, corner_radius=14)
        card.pack(fill="x")
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=18, pady=16)
        return inner

    # ─────────────────────────────────────────
    # DROP ZONE
    # ─────────────────────────────────────────
    def _dz_hover(self, entering: bool):
        if entering:
            self._drop_zone.configure(border_color=ACCENT, fg_color=ACCENT_GLOW)
        else:
            if not self._source_path.get():
                self._drop_zone.configure(border_color=BORDER_C, fg_color=BG_INPUT)

    def _dz_set_file(self, path: str):
        name = Path(path).name
        self._source_path.set(path)
        self._drop_lbl.configure(text=f"✓  {name}", text_color=GREEN)
        self._drop_sub.configure(text=path, text_color=TEXT_DIM)
        self._drop_zone.configure(border_color=GREEN, fg_color=GREEN_DIM)

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
            self._dz_set_file(path)
        else:
            messagebox.showwarning("تنبيه",
                "يرجى إسقاط ملف Excel فقط\n(.xlsx / .xls / .xlsm)")

    # ─────────────────────────────────────────
    # BROWSE
    # ─────────────────────────────────────────
    def _browse_source(self):
        p = filedialog.askopenfilename(
            title="اختر ملف Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")],
        )
        if p:
            self._dz_set_file(p)

    def _browse_output(self):
        p = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if p:
            self._output_dir.set(p)

    # ─────────────────────────────────────────
    # LOG
    # ─────────────────────────────────────────
    def _log(self, msg: str):
        self._log_box.configure(state="normal")

        # Pick tag based on message content
        m = msg.strip()
        if m.startswith("🚀"):
            tag = "start"
        elif m.startswith("✅") or m.startswith("✓") or "بنجاح" in m:
            tag = "ok"
        elif m.startswith("❌") or "خطأ" in m:
            tag = "error"
        elif m.startswith("📁") or m.startswith("   ") and ("C:\\" in m or "/" in m):
            tag = "path"
        elif m.startswith("📝") or m.startswith("📂") or m.startswith("📋"):
            tag = "info"
        elif m.startswith("   ✓"):
            tag = "ok"
        elif m.startswith("   "):
            tag = "step"
        else:
            tag = "info"

        self._log_box._textbox.configure(state="normal")
        self._log_box._textbox.insert("end", msg + "\n", tag)
        self._log_box._textbox.see("end")
        self._log_box._textbox.configure(state="disabled")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_box._textbox.configure(state="normal")
        self._log_box._textbox.delete("1.0", "end")
        self._log_box._textbox.configure(state="disabled")
        self._log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str):
        self._status_lbl.configure(text=f"● {text}", text_color=color)

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
        self._progress.set(0)
        self._set_status("جارٍ المعالجة ...", AMBER)
        self._log_section.show()
        self._clear_log()
        self._log("🚀 بدء المعالجة...")

        def worker():
            try:
                self.after(0, self._progress.set, 0.2)
                raw_df, filtered_df = load_and_filter(
                    path, self._after_6pm.get(),
                    lambda m: self.after(0, self._log, m),
                )
                self.after(0, self._progress.set, 0.65)
                out = build_workbook(
                    raw_df, filtered_df, out_dir,
                    lambda m: self.after(0, self._log, m),
                )
                self.after(0, self._progress.set, 1.0)
                from config import SECTOR_SHEETS
                stats = {s: len(filtered_df[filtered_df["sector"] == s])
                         for s in SECTOR_SHEETS}
                self.after(0, self._on_success, out, stats)
            except Exception as e:
                self.after(0, self._on_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    # ─────────────────────────────────────────
    # CALLBACKS
    # ─────────────────────────────────────────
    def _on_success(self, out_path: str, stats: dict):
        self._last_output = out_path
        self._start_btn.configure(state="normal", text="▶   ابدأ المعالجة")
        self._open_btn.configure(state="normal")
        self._set_status("اكتملت المعالجة بنجاح  ✓", GREEN)
        self._log(f"\n📁 الملف محفوظ في:\n   {out_path}")
        self._stats_section.show()
        self._update_stats(stats)
        messagebox.showinfo("تم بنجاح ✅",
                            f"تم إنشاء التقرير بنجاح!\n\n{out_path}")

    def _on_error(self, msg: str):
        self._start_btn.configure(state="normal", text="▶   ابدأ المعالجة")
        self._progress.set(0)
        self._set_status("حدث خطأ أثناء المعالجة", RED)
        self._log(f"\n❌ خطأ: {msg}")
        messagebox.showerror("خطأ", msg)

    def _open_output(self):
        if self._last_output and os.path.exists(self._last_output):
            if sys.platform == "win32":
                os.startfile(self._last_output)  # type: ignore
            elif sys.platform == "darwin":
                subprocess.call(["open", self._last_output])
            else:
                subprocess.call(["xdg-open", self._last_output])

    # ─────────────────────────────────────────
    # STATS
    # ─────────────────────────────────────────
    def _update_stats(self, stats: dict):
        for w in self._stats_inner.winfo_children():
            w.destroy()
        grid = ctk.CTkFrame(self._stats_inner, fg_color="transparent")
        grid.pack(fill="x")
        for i, (sector, count) in enumerate(stats.items()):
            color = STAT_COLORS[i % len(STAT_COLORS)]
            StatCard(grid, sector, count, color=color).grid(
                row=0, column=i, padx=5, pady=4, sticky="ew")
            grid.columnconfigure(i, weight=1)
        total = sum(stats.values())
        self._total_lbl.configure(text=f"الإجمالي الكلى  :  {total}  ارسالية")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()