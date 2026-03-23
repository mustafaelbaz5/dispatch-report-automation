"""
gui/widgets.py — Reusable CustomTkinter widgets shared across the whole app.
"""

import customtkinter as ctk
from gui.theme import (
    BORDER, BG_HOVER, BG_INPUT,
    G, G_H, T, TEXT_DIM, TEXT_SEC, TEXT_PRI,
    font,
)


class Divider(ctk.CTkFrame):
    """A 1 px horizontal rule."""

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BORDER, height=1, **kw)


class PrimaryBtn(ctk.CTkButton):
    """Solid filled action button."""

    def __init__(self, parent, text: str, cmd,
                 icon: str = "", height: int = 46,
                 color: str = G, hcolor: str | None = None, **kw):
        label = f"{icon}  {text}" if icon else text
        super().__init__(
            parent,
            text=label,
            font=font(13, bold=True),
            fg_color=color,
            hover_color=hcolor or G_H,
            text_color="#FFFFFF",
            height=height,
            corner_radius=10,
            command=cmd,
            **kw,
        )


class GhostBtn(ctk.CTkButton):
    """Transparent button with a colored border and text."""

    def __init__(self, parent, text: str, cmd,
                 accent: str = T, icon: str = "",
                 height: int = 38, **kw):
        label = f"{icon}  {text}" if icon else text
        super().__init__(
            parent,
            text=label,
            font=font(11, bold=True),
            fg_color="transparent",
            hover_color=BG_HOVER,
            text_color=accent,
            height=height,
            corner_radius=8,
            border_width=1,
            border_color=accent,
            command=cmd,
            **kw,
        )


class StatusDot(ctk.CTkFrame):
    """Small colored dot + status label."""

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._dot = ctk.CTkFrame(self, width=7, height=7,
                                 corner_radius=4, fg_color=TEXT_DIM)
        self._dot.pack(side="left", padx=(0, 5))
        self._lbl = ctk.CTkLabel(self, text="جاهز",
                                 font=font(10), text_color=TEXT_DIM)
        self._lbl.pack(side="left")

    def set(self, text: str, color: str) -> None:
        self._dot.configure(fg_color=color)
        self._lbl.configure(text=text, text_color=color)