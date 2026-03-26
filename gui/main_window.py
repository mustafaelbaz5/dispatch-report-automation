import os
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

try:
    from tkinterdnd2 import TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

from config import DEFAULT_SAVE_DIR, SECTOR_SHEETS
from core.processor import load_and_filter
from core.writer import build_workbook

from gui.theme import (
    BG_APP, BG_SURFACE, BORDER,
    AMBER, RED, SUCCESS, TEXT_DIM,
)
from gui.panels.header      import HeaderPanel
from gui.panels.dropzone    import DropzoneCard
from gui.panels.output_row import OutputDirCard
from gui.panels.options_row import OptionsCard
from gui.panels.actions     import ActionsPanel
from gui.panels.log_panel   import LogPanel
from gui.widgets import StatusDot
from gui.dialogs.print_dialog   import PrintDialog
from gui.dialogs.ijmaly_dialog import IjmalyPrintDialog

# 1. وضع الدالة خارج الكلاس لتكون متاحة لكل المشروع
def resource_path(relative_path):
    """ الحصول على المسار الصحيح للملفات المدمجة داخل الـ EXE """
    try:
        base_path = sys._MEIPASS # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class MainWindow(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        if _DND_AVAILABLE:
            try:
                TkinterDnD._require(self)  # type: ignore
            except Exception:
                pass

        self.title("بيان تسليم الارساليات الصادرة")
        
        # 2. استخدام resource_path لتحميل الأيقونة بشكل صحيح
        from PIL import Image, ImageTk
        try:
            logo_path = resource_path("car_logo.png")
            icon_img = Image.open(logo_path)
            icon = ImageTk.PhotoImage(icon_img)
            self.iconphoto(True, icon) # pyright: ignore[reportArgumentType]
        except Exception as e:
            print(f"Icon Load Error: {e}")

        self.configure(fg_color=BG_APP)
        self._center_window()
        self.minsize(580, 600)

        # ── Shared state ──────────────────────────────────────────────────────
        self._source_path   = ctk.StringVar()
        self._output_dir    = ctk.StringVar(value=str(DEFAULT_SAVE_DIR))
        self._after_6pm     = ctk.BooleanVar(value=False)
        self._last_output: str | None = None
        self._sector_totals: dict     = {}

        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _center_window(self) -> None:
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{sw // 2}x{sh - 80}+0+0")
        self.resizable(True, True)

    def _build_ui(self) -> None:
        HeaderPanel(self).pack(fill="x")

        scroll = ctk.CTkScrollableFrame(
            self, fg_color=BG_APP,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color="#64748B")
        scroll.pack(fill="both", expand=True)

        PAD = {"fill": "x", "padx": 14}

        DropzoneCard(scroll, self._source_path).pack(**PAD, pady=(12, 0))
        OutputDirCard(scroll, self._output_dir).pack(**PAD, pady=(8, 0))
        OptionsCard(scroll, self._after_6pm).pack(**PAD, pady=(8, 0))

        self._actions = ActionsPanel(
            scroll,
            on_run=self._run,
            on_open=self._open_output,
            on_print=self._open_print_dialog,
            on_print_ijmaly=self._open_ijmaly_dialog,
            on_reset=self._reset,
        )
        self._actions.pack(**PAD, pady=(10, 0))

        self._build_progress(scroll) # pyright: ignore[reportArgumentType]
        self._log_panel = LogPanel(scroll)

    def _build_progress(self, parent: ctk.CTkFrame) -> None:
        pf = ctk.CTkFrame(parent, fg_color="transparent")
        pf.pack(fill="x", padx=14, pady=(10, 0))

        self._progress = ctk.CTkProgressBar(
            pf, height=4, progress_color=SUCCESS,
            fg_color=BORDER, corner_radius=2)
        self._progress.pack(fill="x", pady=(0, 5))
        self._progress.set(0)

        row = ctk.CTkFrame(pf, fg_color="transparent")
        row.pack(fill="x")

        self._status = StatusDot(row)
        self._status.pack(side="right")

        self._log_toggle_btn = ctk.CTkButton(
            row, text="▾  السجل",
            font=ctk.CTkFont(size=9),
            fg_color="transparent", hover_color="#1C2A3A",
            text_color="#64748B", height=22, width=72,
            border_width=1, border_color=BORDER, corner_radius=6,
            command=self._toggle_log)
        self._log_toggle_btn.pack(side="left")

    def _log(self, msg: str) -> None:
        self._log_panel.log(msg)
        if not self._log_panel.visible:
            self._log_panel.show()
            self._log_toggle_btn.configure(text="▴  السجل")

    def _toggle_log(self) -> None:
        visible = self._log_panel.toggle()
        self._log_toggle_btn.configure(
            text="▴  السجل" if visible else "▾  السجل")

    def _reset(self) -> None:
        for widget in self.winfo_children():
            self._find_and_reset_dropzone(widget)

        self._last_output   = None
        self._sector_totals = {}
        self._actions.set_output_ready(False)
        self._actions.set_processing(False)
        self._progress.set(0)
        self._status.set("جاهز", TEXT_DIM)
        self._log_panel.clear()

    def _find_and_reset_dropzone(self, widget) -> None:
        if isinstance(widget, DropzoneCard):
            widget.reset()
            return
        for child in widget.winfo_children():
            self._find_and_reset_dropzone(child)

    def _open_print_dialog(self) -> None:
        if not self._output_exists():
            return
        PrintDialog(self, self._last_output) # type: ignore

    def _open_ijmaly_dialog(self) -> None:
        if not self._output_exists():
            return
        IjmalyPrintDialog(self, self._last_output, self._sector_totals) # pyright: ignore[reportArgumentType]

    def _output_exists(self) -> bool:
        if not self._last_output or not Path(self._last_output).exists():
            messagebox.showwarning("تنبيه", "يرجى تشغيل المعالجة أولاً.")
            return False
        return True

    def _run(self) -> None:
        path    = self._source_path.get().strip()
        out_dir = self._output_dir.get().strip() or str(DEFAULT_SAVE_DIR)

        if not path:
            messagebox.showwarning("تنبيه", "يرجى اختيار ملف Excel أولاً.")
            return
        if not os.path.exists(path):
            messagebox.showerror("خطأ", f"الملف غير موجود:\n{path}")
            return

        Path(out_dir).mkdir(parents=True, exist_ok=True)
        self._actions.set_processing(True)
        self._actions.set_output_ready(False)
        self._progress.set(0)
        self._status.set("جارٍ المعالجة ...", AMBER)
        self._log_panel.clear()
        self._log("🚀 بدء المعالجة...")

        def worker() -> None:
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
            except Exception as exc:
                self.after(0, self._on_error, str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, out_path: str, sector_totals: dict) -> None:
        self._last_output   = out_path
        self._sector_totals = sector_totals
        self._actions.set_processing(False)
        self._actions.set_output_ready(True)
        self._status.set("اكتملت المعالجة بنجاح ✓", SUCCESS)
        self._log(f"📁 {out_path}")
        messagebox.showinfo("تم بنجاح ✅",
                            f"تم إنشاء التقرير بنجاح!\n\n{out_path}")

    def _on_error(self, msg: str) -> None:
        self._actions.set_processing(False)
        self._progress.set(0)
        self._status.set("حدث خطأ", RED)
        self._log(f"❌ خطأ: {msg}")
        messagebox.showerror("خطأ", msg)

    def _open_output(self) -> None:
        if not self._last_output or not os.path.exists(self._last_output):
            return
        if sys.platform == "win32":
            os.startfile(self._last_output)
        elif sys.platform == "darwin":
            subprocess.call(["open", self._last_output])
        else:
            subprocess.call(["xdg-open", self._last_output])