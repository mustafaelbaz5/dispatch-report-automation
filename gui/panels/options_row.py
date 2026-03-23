"""
gui/panels/options_row.py — "After 6 PM" filter toggle card.
"""

import customtkinter as ctk

from gui.theme import BG_CARD, BORDER, G, TEXT_PRI, SUCCESS, font


class OptionsCard(ctk.CTkFrame):
    """Card with a single toggle: filter shipments after 6 PM only."""

    def __init__(self, parent, after_6pm_var: ctk.BooleanVar, **kw):
        super().__init__(parent, fg_color=BG_CARD, corner_radius=10,
                         border_width=1, border_color=BORDER, **kw)
        self._var = after_6pm_var
        self._build()

    def _build(self) -> None:
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)

        ctk.CTkLabel(inner,
                     text="⏰  تصفية ارساليات ما بعد 6 مساءً فقط",
                     font=font(11), text_color=TEXT_PRI,
                     anchor="e").pack(side="right")

        ctk.CTkSwitch(inner, text="", variable=self._var,
                      progress_color=G, button_color=TEXT_PRI,
                      button_hover_color=SUCCESS,
                      width=46).pack(side="left")