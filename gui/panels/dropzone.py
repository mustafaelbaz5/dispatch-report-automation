"""
gui/panels/dropzone.py — Source file selection card with drag-and-drop support.
"""

from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

try:
    from tkinterdnd2 import DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False

from gui.theme import (
    BG_CARD, BG_INPUT, BG_HOVER, BORDER,
    G, T,
    TEXT_PRI, TEXT_SEC, TEXT_DIM, SUCCESS,
    font,
)
from gui.widgets import GhostBtn


class DropzoneCard(ctk.CTkFrame):
    """Card that holds the source file path (entry + drag-and-drop zone)."""

    def __init__(self, parent, source_path_var: ctk.StringVar, **kw):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=12,
                         border_width=1, border_color=BORDER, **kw)
        self._var = source_path_var
        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=12)

        ctk.CTkLabel(inner, text="ملف المصدر ( Excel )",
                     font=font(11, bold=True), text_color=TEXT_SEC,
                     anchor="e").pack(anchor="e", pady=(0, 8))

        self._build_dropzone(inner)
        self._build_path_row(inner)

    def _build_dropzone(self, parent: ctk.CTkFrame) -> None:
        self._dz = ctk.CTkFrame(parent, fg_color=BG_INPUT, corner_radius=10,
                                border_width=2, border_color=BORDER, height=130)
        self._dz.pack(fill="x", pady=(0, 10))
        self._dz.pack_propagate(False)

        center = ctk.CTkFrame(self._dz, fg_color="transparent")
        center.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(center, text="⬆", font=font(20), text_color=TEXT_DIM).pack()
        self._dz_lbl = ctk.CTkLabel(center,
                                    text="اسحب ملف Excel وأفلته هنا",
                                    font=font(12, bold=True), text_color=TEXT_SEC)
        self._dz_lbl.pack(pady=(2, 0))
        self._dz_sub = ctk.CTkLabel(center, text=".xlsx  ·  .xls  ·  .xlsm",
                                    font=font(9), text_color=TEXT_DIM)
        self._dz_sub.pack()

        if _DND_AVAILABLE:
            try:
                canvas = self._dz._canvas  # type: ignore
                canvas.drop_target_register(DND_FILES)  # type: ignore
                canvas.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore
            except Exception:
                pass

        self._dz.bind("<Enter>", lambda _: self._set_hover(True))
        self._dz.bind("<Leave>", lambda _: self._set_hover(False))

    def _build_path_row(self, parent: ctk.CTkFrame) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkEntry(row, textvariable=self._var,
                     placeholder_text="أو اكتب / الصق المسار هنا ...",
                     font=font(10), fg_color=BG_INPUT,
                     border_color=BORDER, text_color=TEXT_PRI,
                     height=36, justify="right", corner_radius=8
                     ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        GhostBtn(row, text="تصفح", cmd=self._browse,
                 accent=T, height=36, icon="📁").pack(side="left")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def set_file(self, path: str) -> None:
        self._var.set(path)
        self._dz_lbl.configure(text=f"✓  {Path(path).name}", text_color=SUCCESS)
        self._dz_sub.configure(text=path, text_color=TEXT_DIM)
        self._dz.configure(border_color=G, fg_color="#091A10")

    def reset(self) -> None:
        self._var.set("")
        self._dz_lbl.configure(text="اسحب ملف Excel وأفلته هنا", text_color=TEXT_SEC)
        self._dz_sub.configure(text=".xlsx  ·  .xls  ·  .xlsm", text_color=TEXT_DIM)
        self._dz.configure(border_color=BORDER, fg_color=BG_INPUT)

    def _set_hover(self, entering: bool) -> None:
        if entering:
            self._dz.configure(border_color=T, fg_color="#0D1E35")
        elif not self._var.get():
            self._dz.configure(border_color=BORDER, fg_color=BG_INPUT)

    def _on_drop(self, event) -> None:
        path = event.data.strip().strip("{}")
        if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
            self.set_file(path)
        else:
            messagebox.showwarning("تنبيه",
                "يرجى إسقاط ملف Excel فقط\n(.xlsx / .xls / .xlsm)")

    def _browse(self) -> None:
        p = filedialog.askopenfilename(
            title="اختر ملف Excel",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if p:
            self.set_file(p)