"""
gui/panels/header.py — Top header bar, fully centred.

Configurable constants:
    SUBTITLE  — shown in the badge below the main title
    MANAGER   — manager name shown in the credit line
"""

import customtkinter as ctk
from gui.theme import (
    BG_SURFACE,
    G, G_DIM, T, T_DIM, A,
    TEXT_PRI, TEXT_SEC, TEXT_DIM, SUCCESS,
    font,
)

# ── Configurable labels ────────────────────────────────────────────────────────
SUBTITLE = "عمليات البريد السريع بالمنصورة 9001"
MANAGER  = "محمد شعبان"


class HeaderPanel(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BG_SURFACE,
                         corner_radius=0, height=120, **kw)
        self.pack_propagate(False)
        self._build()

    def _build(self) -> None:
        self._build_top_strip()
        self._build_centre_block()
        self._build_bottom_strip()

    # ── Top colour strip ──────────────────────────────────────────────────────

    def _build_top_strip(self) -> None:
        strip = ctk.CTkFrame(self, fg_color="transparent",
                             height=4, corner_radius=0)
        strip.pack(fill="x", side="top")
        strip.pack_propagate(False)
        ctk.CTkFrame(strip, fg_color=G, height=4,
                     corner_radius=0).pack(side="left", fill="both", expand=True)
        ctk.CTkFrame(strip, fg_color=T, height=4,
                     width=60, corner_radius=0).pack(side="left")
        ctk.CTkFrame(strip, fg_color=A, height=4,
                     width=30, corner_radius=0).pack(side="left")

    # ── Centre block ──────────────────────────────────────────────────────────

    def _build_centre_block(self) -> None:
        block = ctk.CTkFrame(self, fg_color="transparent")
        block.place(relx=0.5, rely=0.46, anchor="center")

        # Main title
        ctk.CTkLabel(
            block,
            text="بيان تسليم الارساليات الصادرة",
            font=font(18, bold=True),
            text_color=TEXT_PRI,
            anchor="center",
        ).pack()

        # Subtitle pill
        pill = ctk.CTkFrame(block, fg_color=T_DIM, corner_radius=6, height=20)
        pill.pack(pady=(5, 0))
        pill.pack_propagate(False)
        ctk.CTkLabel(pill, text=SUBTITLE,
                     font=font(9, bold=True),
                     text_color=T).pack(padx=12, pady=2)

        # Manager credit line
        mgr_row = ctk.CTkFrame(block, fg_color="transparent")
        mgr_row.pack(pady=(6, 0))

        ctk.CTkLabel(
            mgr_row,
            text="تتم الإدارة بواسطة رئيس العمليات البريدية المتخصصة"
                 " والمشرف على المركز اللوجيستى",
            font=font(9),
            text_color=TEXT_SEC,
        ).pack(side="right")

        ctk.CTkLabel(mgr_row, text="  :  ",
                     font=font(9), text_color=TEXT_DIM).pack(side="right")

        name_pill = ctk.CTkFrame(mgr_row, fg_color=G_DIM,
                                 corner_radius=5, height=18)
        name_pill.pack(side="right")
        name_pill.pack_propagate(False)
        ctk.CTkLabel(name_pill, text=f"  {MANAGER}  ",
                     font=font(9, bold=True),
                     text_color=SUCCESS).pack(expand=True)

    # ── Bottom colour strip ───────────────────────────────────────────────────

    def _build_bottom_strip(self) -> None:
        strip = ctk.CTkFrame(self, fg_color="transparent",
                             height=2, corner_radius=0)
        strip.pack(fill="x", side="bottom")
        strip.pack_propagate(False)
        ctk.CTkFrame(strip, fg_color=G_DIM, height=2,
                     corner_radius=0).pack(fill="both", expand=True)