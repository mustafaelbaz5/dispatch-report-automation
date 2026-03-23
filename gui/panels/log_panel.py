"""
gui/panels/log_panel.py — Collapsible operation log panel.
"""

import customtkinter as ctk

from gui.theme import (
    BG_SURFACE, BG_INPUT, BG_HOVER, BORDER,
    T, TEXT_SEC, TEXT_DIM,
    RED, SUCCESS,
    font, mono_font,
)
from gui.widgets import Divider


class LogPanel(ctk.CTkFrame):
    """
    Collapsible log panel.

    Usage
    -----
    panel = LogPanel(parent)
    panel.pack(...)          # or let toggle() handle visibility

    panel.log("some message")
    panel.toggle()
    """

    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=BG_SURFACE,
                         corner_radius=10, border_width=1,
                         border_color=BORDER, **kw)
        self._visible = False
        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(8, 4))

        ctk.CTkLabel(hdr, text="◉  سجل العمليات",
                     font=font(9, bold=True), text_color=TEXT_SEC
                     ).pack(side="right")

        ctk.CTkButton(hdr, text="مسح", font=font(8),
                      width=44, height=20,
                      fg_color=BG_INPUT, hover_color=BG_HOVER,
                      text_color=TEXT_DIM, border_width=1,
                      border_color=BORDER, corner_radius=5,
                      command=self.clear).pack(side="left")

        self._box = ctk.CTkTextbox(
            self, font=mono_font(10),
            fg_color="transparent", text_color="#94A3B8",
            border_width=0, height=120, wrap="word")
        self._box.pack(fill="x", padx=10, pady=(0, 8))

        # Configure coloured tags
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
            if fnt:
                kw["font"] = fnt
            self._box._textbox.tag_config(name, **kw)

        self._box.configure(state="disabled")

    # ── Public API ────────────────────────────────────────────────────────────

    @property
    def visible(self) -> bool:
        return self._visible

    def log(self, msg: str) -> None:
        """Append a message with automatic tag detection."""
        m = msg.strip()
        if m.startswith("🚀"):
            tag = "start"
        elif "✅" in m or "بنجاح" in m:
            tag = "ok"
        elif "❌" in m or "خطأ" in m:
            tag = "error"
        elif "📁" in m:
            tag = "path"
        elif m.startswith("   ✓"):
            tag = "ok"
        elif m.startswith("   "):
            tag = "step"
        else:
            tag = "info"

        self._box._textbox.configure(state="normal")
        self._box._textbox.insert("end", msg + "\n", tag)
        self._box._textbox.see("end")
        self._box._textbox.configure(state="disabled")
        self._box.configure(state="disabled")

    def clear(self) -> None:
        self._box._textbox.configure(state="normal")
        self._box._textbox.delete("1.0", "end")
        self._box._textbox.configure(state="disabled")

    def show(self) -> None:
        if not self._visible:
            self.pack(fill="x", padx=14, pady=(8, 10))
            self._visible = True

    def hide(self) -> None:
        if self._visible:
            self.pack_forget()
            self._visible = False

    def toggle(self) -> bool:
        """Toggle visibility. Returns new visible state."""
        if self._visible:
            self.hide()
        else:
            self.show()
        return self._visible