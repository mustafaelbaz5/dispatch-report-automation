"""
gui/theme.py — CTk appearance + shared color/font constants.
"""

import customtkinter as ctk

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Backgrounds ───────────────────────────────────────────────────────────────
BG_APP     = "#0A0D13"
BG_SURFACE = "#111827"
BG_CARD    = "#141E2E"
BG_INPUT   = "#0F1825"
BG_HOVER   = "#1C2A3A"

# ── Borders ───────────────────────────────────────────────────────────────────
BORDER    = "#1E3048"
BORDER_LT = "#2A4060"

# ── Green ─────────────────────────────────────────────────────────────────────
G     = "#10B981"
G_H   = "#059669"
G_DIM = "#052E20"

# ── Teal ──────────────────────────────────────────────────────────────────────
T     = "#38BDF8"
T_H   = "#0EA5E9"
T_DIM = "#0A2540"

# ── Purple ────────────────────────────────────────────────────────────────────
A   = "#818CF8"
A_H = "#6366F1"

# ── Text ──────────────────────────────────────────────────────────────────────
TEXT_PRI  = "#F0F4F8"
TEXT_SEC  = "#64748B"
TEXT_DIM  = "#2D3F54"
TEXT_GOLD = "#F59E0B"

# ── Semantic ──────────────────────────────────────────────────────────────────
RED     = "#F87171"
AMBER   = "#FBBF24"
SUCCESS = "#34D399"


# ── Font helpers ──────────────────────────────────────────────────────────────
def font(size: int = 12, bold: bool = False,
         family: str = "Tajawal") -> ctk.CTkFont:
    return ctk.CTkFont(family=family, size=size,
                       weight="bold" if bold else "normal")


def mono_font(size: int = 10) -> ctk.CTkFont:
    return ctk.CTkFont(family="Consolas", size=size)