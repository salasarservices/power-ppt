"""
Content-zone rendering (brand spec). Body, tables and images are placed at an
explicit (left, top, width) computed by the flow planner, so objects can sit
side-by-side (e.g. a table beside an image) as well as stacked:
  body  — Poppins Regular, slate, 12pt, line-height 1.6, left-aligned, verbatim.
  table — header row blue fill / white Poppins Semi-Bold; body Poppins Regular slate.
  image — source image, aspect-preserved, fit to the given cell width/height.
"""

import io
import math

from PIL import Image as PILImage
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Inches, Pt

from . import geometry as g


# ── height estimators (shared with the flow planner) ─────────────────────────
def estimate_body_height_in(body: str, width: float = g.BODY_WIDTH) -> float:
    """Font-file-free estimate of rendered body height (inches), slightly generous
    so the next block never overlaps the text."""
    char_w_in = 0.5 * g.BODY_SIZE_PT / 72.0          # ~0.5em average glyph advance
    cpl = max(1, int(width / char_w_in))              # chars per line at this width
    line_h_in = g.BODY_SIZE_PT * g.LINE_SPACING / 72.0
    lines = 0.0
    for para in body.split("\n\n"):
        para = para.strip()
        if not para:
            lines += 1
            continue
        lines += max(1, math.ceil(len(para) / cpl)) + 0.4  # + paragraph spacing
    return lines * line_h_in


def estimate_table_height_in(table) -> float:
    return len(list(table.rows or [])) * g.TABLE_ROW_H


def estimate_image_height_in(image_bytes: bytes, width: float = g.BODY_WIDTH) -> float:
    """Height when the image is fit to `width` (aspect-preserved)."""
    try:
        with PILImage.open(io.BytesIO(image_bytes)) as im:
            iw, ih = im.size
    except Exception:
        return 0.0
    if not iw or not ih:
        return 0.0
    return width * (ih / iw)


# ── renderers (position + width supplied by the flow planner) ────────────────
def render_body(slide, body: str, top: float, height: float,
                left: float = g.BODY_LEFT, width: float = g.BODY_WIDTH) -> None:
    if not body or not body.strip():
        return
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE

    first = True
    for para_text in body.split("\n\n"):
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.text = para_text.strip()
        para.line_spacing = g.LINE_SPACING
        para.font.name = g.FONT_BODY
        para.font.size = Pt(g.BODY_SIZE_PT)
        para.font.color.rgb = g.SLATE
        for run in para.runs:
            run.font.name = g.FONT_BODY
            run.font.size = Pt(g.BODY_SIZE_PT)
            run.font.color.rgb = g.SLATE


def render_table(slide, table, top: float,
                 left: float = g.BODY_LEFT, width: float = g.BODY_WIDTH) -> float:
    rows = list(table.rows or [])
    if not rows:
        return 0.0
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    if n_cols == 0:
        return 0.0

    height = n_rows * g.TABLE_ROW_H
    shape = slide.shapes.add_table(
        n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tbl = shape.table
    for r_idx, row in enumerate(rows):
        for c_idx in range(n_cols):
            val = row[c_idx] if c_idx < len(row) else ""
            cell = tbl.cell(r_idx, c_idx)
            cell.text = str(val)
            for para in cell.text_frame.paragraphs:
                para.font.name = g.FONT_BODY
                if r_idx == 0:
                    para.font.bold = True
                    para.font.size = Pt(g.TABLE_HEAD_PT)
                    para.font.color.rgb = g.WHITE
                else:
                    para.font.size = Pt(g.TABLE_BODY_PT)
                    para.font.color.rgb = g.SLATE
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = g.BLUE
    return height


def render_image(slide, image_bytes: bytes, top: float, zone_bottom: float,
                 left: float = g.BODY_LEFT, width: float = g.BODY_WIDTH) -> float:
    """Fit an image into a cell of `width`, aspect-preserved, clamped to the
    remaining zone height. Returns the vertical space used (inches)."""
    avail_h = zone_bottom - top
    if avail_h < 0.3:
        return 0.0
    try:
        with PILImage.open(io.BytesIO(image_bytes)) as im:
            iw, ih = im.size
    except Exception:
        return 0.0
    if not iw or not ih:
        return 0.0

    aspect = iw / ih
    w = width
    h = w / aspect
    if h > avail_h:                          # too tall -> clamp, keep aspect
        h = avail_h
        w = h * aspect
    slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(left), Inches(top),
                             Inches(w), Inches(h))
    return h


# ── placement renderer ───────────────────────────────────────────────────────
def render_placements(slide, placements) -> None:
    """Render pre-positioned placements. Each is a dict with keys
    kind ('body'|'table'|'image'), payload, left, top, width (+ height for body)."""
    for p in placements:
        kind = p["kind"]
        if kind == "body":
            render_body(slide, p["payload"], p["top"], p["height"], p["left"], p["width"])
        elif kind == "table":
            render_table(slide, p["payload"], p["top"], p["left"], p["width"])
        elif kind == "image":
            render_image(slide, p["payload"], p["top"], g.BODY_BOTTOM, p["left"], p["width"])
