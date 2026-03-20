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
# THEME
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_DARK   = "#0B1120"
BG_CARD   = "#111827"
BG_INPUT  = "#1F2937"
ACCENT    = "#2563EB"
ACCENT_H  = "#1D4ED8"
GREEN     = "#16A34A"
GREEN_H   = "#15803D"
BORDER_C  = "#374151"
TEXT_PRI  = "#F1F5F9"
TEXT_SEC  = "#94A3B8"
TEXT_DIM  = "#4B5563"
RED       = "#DC2626"


def _font(size=16, bold=False):
    return ctk.CTkFont(family="Tajawal", size=size, weight="bold" if bold else "normal")


# ─────────────────────────────────────────────
# WIDGETS
# ─────────────────────────────────────────────
class SectionLabel(ctk.CTkLabel):
    def __init__(self, parent, text, **kw):
        super().__init__(parent, text=text,
                         font=_font(11),
                         text_color=TEXT_SEC,
                         anchor="e", justify="right", **kw)


class InputRow(ctk.CTkFrame):
    """Label + Entry + optional Button in one row."""
    def __init__(self, parent, label, var, placeholder, btn_text=None,
                 btn_cmd=None, btn_color=ACCENT, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        SectionLabel(self, label).pack(anchor="e", pady=(0, 4))
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x")
        self.entry = ctk.CTkEntry(
            row, textvariable=var,
            placeholder_text=placeholder,
            font=_font(11),
            fg_color=BG_INPUT, border_color=BORDER_C,
            text_color=TEXT_PRI, height=40, justify="right",
        )
        self.entry.pack(side="left", fill="x", expand=True,
                        padx=(0, 8) if btn_text else 0)
        if btn_text:
            ctk.CTkButton(
                row, text=btn_text, font=_font(11, bold=True),
                fg_color=btn_color, hover_color=ACCENT_H,
                width=95, height=40, command=btn_cmd,
            ).pack(side="left")


class StatCard(ctk.CTkFrame):
    """Small card showing a sector name + count."""
    def __init__(self, parent, sector, count, **kw):
        super().__init__(parent, fg_color=BG_INPUT,
                         corner_radius=8, **kw)
        ctk.CTkLabel(self, text=sector, font=_font(11, bold=True),
                     text_color=TEXT_PRI).pack(pady=(8, 2))
        ctk.CTkLabel(self, text=str(count), font=_font(18, bold=True),
                     text_color="#60A5FA").pack(pady=(0, 8))


# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Patch DnD onto the underlying Tk window if tkinterdnd2 is available
        if _DND_AVAILABLE:
            try:
                TkinterDnD._require(self) # type: ignore
            except Exception:
                pass
        self.title("بيان تسليم الارساليات الصادرة")
        self.geometry("780x700")
        self.minsize(700, 620)
        self.configure(fg_color=BG_DARK)

        self._source_path = ctk.StringVar()
        self._output_dir  = ctk.StringVar(value=str(DEFAULT_SAVE_DIR))
        self._after_6pm   = ctk.BooleanVar(value=False)
        self._last_output = None   # path of last saved file

        self._build_ui()

    # ─────────────────────────────────────────
    # UI BUILD
    # ─────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_body()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color="#0F2044", corner_radius=0, height=80)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        inner = ctk.CTkFrame(hdr, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(inner,
                     text="📦  بيان تسليم الارساليات الصادرة",
                     font=_font(20, bold=True),
                     text_color=TEXT_PRI).pack()
        ctk.CTkLabel(inner,
                     text="المركز اللوجيستى بالمنصورة 9900",
                     font=_font(11),
                     text_color=TEXT_SEC).pack()

    def _build_body(self):
        scroll = ctk.CTkScrollableFrame(self, fg_color=BG_DARK, scrollbar_button_color=BORDER_C)
        scroll.pack(fill="both", expand=True, padx=0, pady=0)

        # ── Files card ──
        self._files_card = self._card(scroll, "📁  الملفات")

        # Drop zone label
        SectionLabel(self._files_card, "ملف المصدر (Excel):").pack(anchor="e", pady=(0, 4))
        self._drop_zone = ctk.CTkFrame(
            self._files_card,
            fg_color=BG_INPUT, corner_radius=10,
            border_width=2, border_color=BORDER_C,
            height=72,
        )
        self._drop_zone.pack(fill="x", pady=(0, 6))
        self._drop_zone.pack_propagate(False)
        self._drop_lbl = ctk.CTkLabel(
            self._drop_zone,
            text="⬇   اسحب ملف Excel وأفلته هنا  أو اضغط تصفح",
            font=_font(12), text_color=TEXT_DIM,
        )
        self._drop_lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Enable DnD on the drop zone if available
        if _DND_AVAILABLE:
            try:
                # Use the underlying tk widget for DnD registration
                dz = self._drop_zone._canvas  # type: ignore
                dz.drop_target_register(DND_FILES)  # type: ignore
                dz.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore
            except Exception:
                pass
        self._drop_zone.bind("<Enter>",
            lambda e: self._drop_zone.configure(border_color=ACCENT))
        self._drop_zone.bind("<Leave>",
            lambda e: self._drop_zone.configure(border_color=BORDER_C))

        # Browse button row below drop zone
        browse_row = ctk.CTkFrame(self._files_card, fg_color="transparent")
        browse_row.pack(fill="x", pady=(0, 10))
        ctk.CTkEntry(browse_row, textvariable=self._source_path,
                     placeholder_text="أو اكتب المسار مباشرة...",
                     font=_font(11), fg_color=BG_INPUT, border_color=BORDER_C,
                     text_color=TEXT_PRI, height=36, justify="right",
                     ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(browse_row, text="📁 تصفح",
                      font=_font(11, bold=True),
                      fg_color=ACCENT, hover_color=ACCENT_H,
                      width=95, height=36,
                      command=self._browse_source).pack(side="left")

        InputRow(self._files_card, "مجلد الحفظ:",
                 self._output_dir, "مجلد حفظ التقارير...",
                 btn_text="تصفح", btn_cmd=self._browse_output,
                 btn_color="#475569").pack(fill="x")

        # ── Options card ──
        opt_card = self._card(scroll, "⚙️  خيارات المعالجة")
        opt_row  = ctk.CTkFrame(opt_card, fg_color="transparent")
        opt_row.pack(fill="x")
        ctk.CTkSwitch(opt_row,
                      text="تضمين الارساليات بعد 6 مساءً فقط",
                      variable=self._after_6pm,
                      font=_font(12),
                      text_color=TEXT_PRI,
                      progress_color=ACCENT,
                      ).pack(side="right")

        # ── Action buttons ──
        btn_card = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_card.pack(fill="x", padx=24, pady=(0, 8))

        self._start_btn = ctk.CTkButton(
            btn_card, text="▶   ابدأ المعالجة",
            font=_font(14, bold=True),
            fg_color=GREEN, hover_color=GREEN_H,
            height=48, corner_radius=10,
            command=self._run,
        )
        self._start_btn.pack(fill="x", pady=(0, 8))

        self._open_btn = ctk.CTkButton(
            btn_card, text="📂   فتح الملف الناتج",
            font=_font(12),
            fg_color="#1E3A5F", hover_color="#274d7a",
            height=38, corner_radius=8,
            command=self._open_output,
            state="disabled",
        )
        self._open_btn.pack(fill="x")

        # ── Progress bar ──
        self._progress = ctk.CTkProgressBar(scroll, height=6,
                                             progress_color=ACCENT,
                                             fg_color=BORDER_C)
        self._progress.pack(fill="x", padx=24, pady=(10, 4))
        self._progress.set(0)

        # ── Status label ──
        self._status_lbl = ctk.CTkLabel(scroll, text="● جاهز",
                                         font=_font(11),
                                         text_color=TEXT_DIM,
                                         anchor="e")
        self._status_lbl.pack(anchor="e", padx=26, pady=(0, 6))

        # ── Log card ──
        log_card = self._card(scroll, "📋  السجل")
        self._log_box = ctk.CTkTextbox(
            log_card,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#0B1120", text_color="#7DD3FC",
            border_color=BORDER_C, border_width=1,
            height=160, wrap="word",
        )
        self._log_box.pack(fill="x")
        self._log_box.configure(state="disabled")

        btn_row = ctk.CTkFrame(log_card, fg_color="transparent")
        btn_row.pack(fill="x", pady=(8, 0))
        ctk.CTkButton(btn_row, text="مسح السجل",
                      font=_font(10), width=90, height=28,
                      fg_color=BORDER_C, hover_color="#4B5563",
                      command=self._clear_log).pack(side="left")

        # ── Stats card (hidden until first run) ──
        self._stats_card_frame = self._card(scroll, "📊  ملخص النتائج")
        self._stats_inner = ctk.CTkFrame(self._stats_card_frame, fg_color="transparent")
        self._stats_inner.pack(fill="x")
        ctk.CTkLabel(self._stats_card_frame,
                     text="ستظهر هنا إحصائيات كل قطاع بعد المعالجة.",
                     font=_font(11), text_color=TEXT_DIM).pack()
        self._stats_placeholder = self._stats_card_frame.winfo_children()[-1]

    def _card(self, parent, title: str) -> ctk.CTkFrame:
        """Create a titled card section."""
        wrapper = ctk.CTkFrame(parent, fg_color="transparent")
        wrapper.pack(fill="x", padx=24, pady=(0, 12))
        ctk.CTkLabel(wrapper, text=title,
                     font=_font(13, bold=True),
                     text_color=TEXT_SEC).pack(anchor="e", pady=(0, 6))
        card = ctk.CTkFrame(wrapper, fg_color=BG_CARD, corner_radius=12)
        card.pack(fill="x")
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)
        return inner

    # ─────────────────────────────────────────
    # BROWSE
    # ─────────────────────────────────────────
    def _on_drop(self, event):
        """Handle drag-and-drop file onto the drop zone."""
        path = event.data.strip().strip("{}")   # remove braces Windows adds
        if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
            self._source_path.set(path)
            self._drop_lbl.configure(
                text=f"✓  {Path(path).name}", text_color="#4ADE80")
            self._drop_zone.configure(border_color="#16A34A")
        else:
            messagebox.showwarning("تنبيه", "يرجى إسقاط ملف Excel فقط (.xlsx / .xls / .xlsm)")

    def _browse_source(self):
        p = filedialog.askopenfilename(
            title="اختر ملف Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if p:
            self._source_path.set(p)
            self._drop_lbl.configure(
                text=f"✓  {Path(p).name}", text_color="#4ADE80")
            self._drop_zone.configure(border_color="#16A34A")

    def _browse_output(self):
        p = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if p:
            self._output_dir.set(p)

    # ─────────────────────────────────────────
    # LOG
    # ─────────────────────────────────────────
    def _log(self, msg: str):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    def _set_status(self, text: str, color: str):
        self._status_lbl.configure(text=text, text_color=color)

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
        if out_dir:
            Path(out_dir).mkdir(parents=True, exist_ok=True)

        # Reset UI
        self._start_btn.configure(state="disabled", text="⏳  جارٍ المعالجة...")
        self._open_btn.configure(state="disabled")
        self._progress.set(0)
        self._set_status("● جارٍ المعالجة...", "#FBBF24")
        self._clear_log()
        self._log("🚀 بدء المعالجة...")
        self._clear_stats()

        def worker():
            try:
                # Step 1 — load & filter  (progress 0 → 0.5)
                self.after(0, self._progress.set, 0.2)
                raw_df, filtered_df = load_and_filter(
                    path, self._after_6pm.get(),
                    lambda m: self.after(0, self._log, m)
                )

                # Step 2 — build Excel  (progress 0.5 → 1.0)
                self.after(0, self._progress.set, 0.6)
                out = build_workbook(
                    raw_df, filtered_df, out_dir,
                    lambda m: self.after(0, self._log, m)
                )
                self.after(0, self._progress.set, 1.0)

                # Build stats
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
        self._set_status("● اكتملت المعالجة بنجاح ✓", "#4ADE80")
        self._update_stats(stats)
        self._log(f"\n📁 الملف محفوظ في:\n   {out_path}")
        messagebox.showinfo("تم بنجاح ✅",
                            f"تم إنشاء التقرير بنجاح!\n\n{out_path}")

    def _on_error(self, msg: str):
        self._start_btn.configure(state="normal", text="▶   ابدأ المعالجة")
        self._progress.set(0)
        self._set_status("● حدث خطأ", RED)
        self._log(f"\n❌ خطأ: {msg}")
        messagebox.showerror("خطأ", msg)

    def _open_output(self):
        if self._last_output and os.path.exists(self._last_output):
            if sys.platform == "win32":
                os.startfile(self._last_output)
            elif sys.platform == "darwin":
                subprocess.call(["open", self._last_output])
            else:
                subprocess.call(["xdg-open", self._last_output])

    # ─────────────────────────────────────────
    # STATS
    # ─────────────────────────────────────────
    def _clear_stats(self):
        for w in self._stats_inner.winfo_children():
            w.destroy()

    def _update_stats(self, stats: dict):
        self._clear_stats()
        if hasattr(self, "_stats_placeholder"):
            try:
                self._stats_placeholder.destroy()
            except Exception:
                pass

        grid = ctk.CTkFrame(self._stats_inner, fg_color="transparent")
        grid.pack(fill="x")
        for i, (sector, count) in enumerate(stats.items()):
            StatCard(grid, sector, count).grid(
                row=0, column=i, padx=4, pady=4, sticky="ew"
            )
            grid.columnconfigure(i, weight=1)

        total = sum(stats.values())
        ctk.CTkLabel(self._stats_inner,
                     text=f"الإجمالي الكلى :  {total} ارسالية",
                     font=_font(13, bold=True),
                     text_color="#60A5FA").pack(pady=(10, 0))


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    App().mainloop()