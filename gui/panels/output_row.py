"""
gui/panels/output_row.py — Save directory selection card.
"""

from tkinter import filedialog

import customtkinter as ctk
from gui.theme import BG_CARD, BG_INPUT, BORDER, TEXT_PRI, TEXT_SEC, font
from gui.widgets import GhostBtn


class OutputDirCard(ctk.CTkFrame):
    def __init__(self, parent, output_dir_var: ctk.StringVar, **kw):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=10,
                         border_width=1, border_color=BORDER, **kw)
        self._var = output_dir_var
        self._build()

    def _build(self) -> None:
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(inner, text="💾  مجلد الحفظ",
                     font=font(11, bold=True), text_color=TEXT_SEC,
                     anchor="e").pack(anchor="e", pady=(0, 6))

        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")
        ctk.CTkEntry(row, textvariable=self._var,
                     font=font(10), fg_color=BG_INPUT,
                     border_color=BORDER, text_color=TEXT_PRI,
                     height=36, justify="right", corner_radius=8
                     ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        GhostBtn(row, text="تصفح", cmd=self._browse,
                 accent=TEXT_SEC, height=36, icon="📂").pack(side="left")

    def _browse(self) -> None:
        p = filedialog.askdirectory(title="اختر مجلد الحفظ")
        if p:
            self._var.set(p)