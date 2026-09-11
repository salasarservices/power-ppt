"""
Content-zone rendering (brand spec). Body, tables and images are stacked
top-to-bottom in the white content zone with a shared y-cursor, so a slide can
carry text AND a table AND images together without overlap:
  body  — Poppins Regular, slate, 12pt, line-height 1.6, left-aligned, verbatim.
  table — header row blue fill / white Poppins Semi-Bold; body Poppins Regular slate.
  image — source image, aspect-preserved, auto-fit to the content width/height.
"""

import io
import math

from PIL import Image as PILImage
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Emu, Inches, Pt

from . import geometry as g


def _estimate_body_height_in(body: str) -> float:
    """Font-file-free estimate of rendered body height (inches). Slightly generous
    so the next stacked block never overlaps the text."""
    char_w_in = 0.5 * g.BODY_SIZE_PT / 72.0          # ~0.5em average glyph advance
    cpl = max(1, int(g.BODY_WIDTH / char_w_in))       # chars per line
    line_h_in = g.BODY_SIZE_PT * g.LINE_SPACING / 72.0
    lines = 0.0
    for para in body.split("\n\n"):
        para = para.strip()
        if not para:
            lines += 1
            continue
        lines += max(1, math.ceil(len(para) / cpl)) + 0.4  # + paragraph spacing
    return lines * line_h_in


def render_body(slide, body: str, top: float, height: float) -> None:
    if not body or not body.strip():
        return
    box = slide.shapes.add_textbox(
        Inches(g.BODY_LEFT), Inches(top), Inches(g.BODY_WIDTH), Inches(height)
    )
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


def render_table(slide, table, top: float) -> float:
    """Render a table at `top`; return the vertical space used (inches)."""
    rows = list(table.rows or [])
    if not rows:
        return 0.0
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=0)
    if n_cols == 0:
        return 0.0

    height = n_rows * g.TABLE_ROW_H
    shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(g.BODY_LEFT), Inches(top),
        Inches(g.BODY_WIDTH), Inches(height),
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


def render_image(slide, image_bytes: bytes, top: float, zone_bottom: float) -> float:
    """Place a source image at `top`, aspect-preserved, fit to the content width
    and the remaining height. Returns the vertical space used (inches); 0 if no room."""
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
    width = g.BODY_WIDTH
    height = width / aspect
    if height > avail_h:                     # too tall -> clamp to remaining height
        height = avail_h
        width = height * aspect
    slide.shapes.add_picture(
        io.BytesIO(image_bytes),
        Inches(g.BODY_LEFT), Inches(top),
        Inches(width), Inches(height),
    )
    return height


# ── height estimators (shared with the flow planner) ─────────────────────────
def estimate_body_height_in(body: str) -> float:
    return _estimate_body_height_in(body)


def estimate_table_height_in(table) -> float:
    return len(list(table.rows or [])) * g.TABLE_ROW_H


def estimate_image_height_in(image_bytes: bytes) -> float:
    """Height when the image is fit to the content width (aspect-preserved)."""
    try:
        with PILImage.open(io.BytesIO(image_bytes)) as im:
            iw, ih = im.size
    except Exception:
        return 0.0
    if not iw or not ih:
        return 0.0
    return g.BODY_WIDTH * (ih / iw)


# ── block renderer (blocks are pre-sized to fit by the flow planner) ──────────
def render_blocks(slide, blocks) -> None:
    """Render a slide's flowed blocks top-to-bottom. Each block is
    ("body", str) | ("table", Table) | ("image", bytes)."""
    y = g.BODY_TOP
    for kind, payload in blocks:
        if kind == "body":
            h = min(_estimate_body_height_in(payload), g.BODY_BOTTOM - y)
            render_body(slide, payload, y, h)
            y += h + g.CONTENT_GAP
        elif kind == "table":
            used = render_table(slide, payload, y)
            if used > 0:
                y += used + g.CONTENT_GAP
        elif kind == "image":
            used = render_image(slide, payload, y, g.BODY_BOTTOM)
            if used > 0:
                y += used + g.CONTENT_GAP
