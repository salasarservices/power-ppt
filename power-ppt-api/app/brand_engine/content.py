"""
Body-text and table rendering into the white content zone (brand spec):
  body  — Poppins Regular, slate, 14pt, line-height 1.6, left-aligned, content verbatim.
  table — header row blue fill / white Poppins Semi-Bold; body Poppins Regular slate.
"""

from pptx.util import Inches, Pt

from . import geometry as g


def render_body(slide, body: str) -> None:
    if not body or not body.strip():
        return

    box = slide.shapes.add_textbox(
        Inches(g.BODY_LEFT), Inches(g.BODY_TOP),
        Inches(g.BODY_WIDTH), Inches(g.BODY_HEIGHT),
    )
    tf = box.text_frame
    tf.word_wrap = True

    paragraphs = [p for p in body.split("\n\n")]
    first = True
    for para_text in paragraphs:
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


def render_table(slide, table) -> None:
    rows = list(table.rows or [])
    if not rows:
        return
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    if n_rows == 0 or n_cols == 0:
        return

    row_h = Inches(0.34)
    shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(g.BODY_LEFT), Inches(g.BODY_TOP),
        Inches(g.BODY_WIDTH), row_h * n_rows,
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
