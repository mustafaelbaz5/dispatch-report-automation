"""
gui/panels/actions.py — Action buttons panel (start, open, print, reset).
"""

import customtkinter as ctk

from gui.theme import A, A_H, T, TEXT_GOLD, TEXT_SEC, G, font
from gui.widgets import GhostBtn, PrimaryBtn


class ActionsPanel(ctk.CTkFrame):
    """
    Holds all action buttons.

    Callbacks are injected by MainWindow so this panel stays free
    of any business logic.
    """

    def __init__(self, parent,
                 on_run,
                 on_open,
                 on_print,
                 on_print_ijmaly,
                 on_reset,
                 **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self._on_run          = on_run
        self._on_open         = on_open
        self._on_print        = on_print
        self._on_print_ijmaly = on_print_ijmaly
        self._on_reset        = on_reset
        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        self._start_btn = PrimaryBtn(self, text="ابدأ المعالجة",
                                     cmd=self._on_run,
                                     icon="▶", height=48)
        self._start_btn.pack(fill="x", pady=(0, 8))

        self._build_secondary_row()

        GhostBtn(self, text="تصفير", cmd=self._on_reset,
                 accent=TEXT_SEC, height=34, icon="↺").pack(fill="x")

    def _build_secondary_row(self) -> None:
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", pady=(0, 6))
        row.columnconfigure(0, weight=1)
        row.columnconfigure(1, weight=1)
        row.columnconfigure(2, weight=1)

        self._open_btn = GhostBtn(row, text="فتح الملف",
                                   cmd=self._on_open,
                                   accent=T, height=38, icon="📂")
        self._open_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self._print_btn = GhostBtn(row, text="طباعة الشيتات",
                                    cmd=self._on_print,
                                    accent=A, height=38, icon="🖨")
        self._print_btn.grid(row=0, column=1, sticky="ew", padx=4)

        self._print_ijmaly_btn = GhostBtn(row, text="طباعة الاجمالى",
                                           cmd=self._on_print_ijmaly,
                                           accent=TEXT_GOLD, height=38, icon="📊")
        self._print_ijmaly_btn.grid(row=0, column=2, sticky="ew", padx=(4, 0))

        # Start disabled — enabled after successful processing
        self.set_output_ready(False)

    # ── Public API ────────────────────────────────────────────────────────────

    def set_processing(self, active: bool) -> None:
        """Disable/re-enable the start button while worker thread runs."""
        if active:
            self._start_btn.configure(state="disabled", text="⏳  جارٍ المعالجة...")
        else:
            self._start_btn.configure(state="normal", text="▶  ابدأ المعالجة")

    def set_output_ready(self, ready: bool) -> None:
        """Enable/disable output buttons depending on whether a file exists."""
        state = "normal" if ready else "disabled"
        self._open_btn.configure(state=state)
        self._print_btn.configure(state=state)
        self._print_ijmaly_btn.configure(state=state)