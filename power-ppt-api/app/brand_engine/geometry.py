"""
Geometry and brand constants for the 2026 Salasar template.

All positions in inches (convert with pptx.util.Inches). Values were measured from
the authorised template `templates/2026/Salasar_Corporate_Blank.pptx`:
  slide 13.33 x 7.50 ; logo top-right from x 10.45 ; footer bar top at 7.03.
See PHASE1-BUILD-PLAN.md section 3.
"""

from pptx.dml.color import RGBColor

# ── Slide ───────────────────────────────────────────────────────────────────
SLIDE_W_IN = 13.33
SLIDE_H_IN = 7.50

# ── Margins / no-go zones ───────────────────────────────────────────────────
MARGIN_L = 0.75
MARGIN_R = 0.75              # right content edge = 12.58
LOGO_LEFT_EDGE = 10.45      # heading must stay left of this

# ── Heading (top-left, clears the logo) ─────────────────────────────────────
HEAD_LEFT = 0.75
HEAD_TOP = 0.55
HEAD_WIDTH = 9.50           # 0.75 -> 10.25, clears the logo
HEAD_HEIGHT = 0.90

# Green rule under the first word. Part of the documented brand slide spec, but
# turned OFF per review (09 Sep 2026). Kept as a reversible toggle.
SHOW_GREEN_RULE = False
RULE_GAP = 0.08             # gap below heading text to the green rule
RULE_HEIGHT = 0.055
RULE_MIN_W = 0.80           # fallback rule width

# ── Content zone (body / tables) — below heading, above footer bar (top 7.03) ─
BODY_LEFT = 0.75
BODY_TOP = 1.55
BODY_WIDTH = 11.83          # 0.75 -> 12.58
BODY_BOTTOM = 6.83          # footer top 7.03 - 0.20 margin
BODY_HEIGHT = BODY_BOTTOM - BODY_TOP   # 5.28
CONTENT_GAP = 0.15          # vertical gap between stacked blocks (body/table/image)
TABLE_ROW_H = 0.34          # nominal table row height (inches)

# ── Colours (brand) ─────────────────────────────────────────────────────────
BLUE = RGBColor(0x1A, 0x3A, 0x8F)
GREEN = RGBColor(0x7A, 0xC1, 0x43)
MID_BLUE = RGBColor(0x00, 0x70, 0xC0)
SLATE = RGBColor(0x4D, 0x4D, 0x4D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ── Fonts ───────────────────────────────────────────────────────────────────
FONT_HEADING_SB = "Poppins SemiBold"   # brand: Semi-Bold headings
FONT_BODY = "Poppins"                  # brand: Regular body

# ── Sizes ───────────────────────────────────────────────────────────────────
HEAD_SIZE_PT = 20
HEAD_MIN_PT = 12
BODY_SIZE_PT = 12
TABLE_HEAD_PT = 11
TABLE_BODY_PT = 10
LINE_SPACING = 1.6                     # brand: line-height 1.6

# ── Pagination ──────────────────────────────────────────────────────────────
# Chars-per-page heuristic sized to the content zone at 14pt Poppins.
DEFAULT_CHARS_PER_PAGE = 1800
CONTINUATION_SUFFIX = " (CONTD...)"
