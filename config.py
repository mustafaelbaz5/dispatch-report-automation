"""
config.py — All constants, column indices, colors, and styles.
Edit this file to change any setting without touching logic code.
"""

from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────────────────────
# SOURCE FILE — COLUMN INDICES (0-based, header=1)
# ─────────────────────────────────────────────
COL_DEST_OFFICE = 4   # "3000-Alexandria Traffic Center"
COL_DISPATCH_NO = 6   # Shipment / Dispatch number
COL_DATETIME    = 8   # Dispatch Date
COL_STATUS      = 9   # Status: "Closed" or "FullyReceived"
COL_TOTAL_ITEMS = 12  # Total Items
COL_WEIGHT      = 14  # Bag weight (KG)

# ─────────────────────────────────────────────
# FILTER SETTINGS
# ─────────────────────────────────────────────
VALID_STATUSES = {"closed", "fullyreceived"}

# ─────────────────────────────────────────────
# REPORT SETTINGS
# ─────────────────────────────────────────────
CENTER_TITLE  = "المركز اللوجيستى بالمنصورة 9900"
SECTOR_SHEETS = ["بحرى", "قبلى", "رمسيس", "طنطا", "المنصورة"]
OUT_HEADERS   = ["م", "الكود", "اسم مركز الحركة", "رقم الارسالية", "الوزن", "ملاحظات"]
COL_WIDTHS    = [6, 12, 28, 20, 12, 35]

# ─────────────────────────────────────────────
# DEFAULT SAVE PATH
# ─────────────────────────────────────────────
DEFAULT_SAVE_DIR = Path.home() / "Desktop" / "التقارير"

# ─────────────────────────────────────────────
# EXCEL COLOR PALETTE
# ─────────────────────────────────────────────
C_NAVY      = "1A3557"   # deep navy   — main headers
C_GOLD_FILL = "FFF3CD"   # light gold  — sub-header row
C_LIGHT_BG  = "F0F4FA"   # light blue  — alternate data rows
C_WHITE     = "FFFFFF"
C_DARK_TEXT = "1A1A2E"
C_SUM_BG    = "E8F0FE"   # soft blue   — summary row
C_SIG_BG    = "FFF8E7"   # warm cream  — signature row

# ─────────────────────────────────────────────
# EXCEL FILLS
# ─────────────────────────────────────────────
NAVY_FILL    = PatternFill("solid", start_color=C_NAVY,     end_color=C_NAVY)
GOLD_BG_FILL = PatternFill("solid", start_color=C_GOLD_FILL,end_color=C_GOLD_FILL)
ALT_FILL     = PatternFill("solid", start_color=C_LIGHT_BG, end_color=C_LIGHT_BG)
WHITE_FILL   = PatternFill("solid", start_color=C_WHITE,    end_color=C_WHITE)
SUM_FILL     = PatternFill("solid", start_color=C_SUM_BG,   end_color=C_SUM_BG)
SIG_FILL     = PatternFill("solid", start_color=C_SIG_BG,   end_color=C_SIG_BG)

# ─────────────────────────────────────────────
# EXCEL FONTS
# ─────────────────────────────────────────────
HDR_FONT   = Font(bold=True,   color=C_WHITE,     name="Arial", size=12)
BIG_FONT   = Font(bold=True,   color=C_WHITE,     name="Arial", size=15)
GOLD_FONT  = Font(bold=True,   color=C_NAVY,      name="Arial", size=12)
SUM_FONT   = Font(bold=True,   color=C_NAVY,      name="Arial", size=11)
DATA_FONT  = Font(             color=C_DARK_TEXT, name="Arial", size=10)
SIG_FONT   = Font(italic=True, color="555555",    name="Arial", size=10)
DATE_FONT  = Font(bold=True,   color=C_NAVY,      name="Arial", size=11)

# ─────────────────────────────────────────────
# EXCEL BORDERS
# ─────────────────────────────────────────────
THIN         = Side(style="thin",   color="B0BEC5")
THICK        = Side(style="medium", color=C_NAVY)
BORDER       = Border(left=THIN,  right=THIN,  top=THIN,  bottom=THIN)
THICK_BORDER = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
BOT_THICK    = Border(left=THIN,  right=THIN,  top=THIN,  bottom=THICK)
TOP_THICK    = Border(left=THICK, right=THICK, top=THICK, bottom=THIN)