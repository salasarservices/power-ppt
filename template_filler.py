# template_filler.py
import io
from typing import List, Optional, Dict

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# Constants for defaults
TITLE_FONT_NAME = "Poppins"
BODY_FONT_NAME = "Poppins"
TITLE_PX = 20
BODY_PX = 12
TITLE_PT = Pt(TITLE_PX * 0.75)
BODY_PT = Pt(BODY_PX * 0.75)
TITLE_COLOR = RGBColor(0x2d, 0x44, 0x8d)
BODY_COLOR = RGBColor(0x00, 0x00, 0x00)


def find_layout_index_with_title_and_body(prs: Presentation) -> int:
    """
    Find a slide layout index with both title and body placeholders. Fallback to 0.
    """
    from pptx.enum.shapes import PP_PLACEHOLDER

    for idx, layout in enumerate(prs.slide_layouts):
        has_title = False
        has_body = False
        for ph in layout.placeholders:
            try:
                if ph.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                    has_title = True
                if ph.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT):
                    has_body = True
            except Exception:
                pass
        if has_title and has_body:
            return idx
    return 0


def _set_paragraph_font(paragraph, name: str, size: Pt, color: RGBColor, bold: bool = False):
    for run in paragraph.runs:
        run.font.name = name
        run.font.size = size
        run.font.color.rgb = color
        run.font.bold = bold
    if not paragraph.runs:
        paragraph.font.name = name
        paragraph.font.size = size
        paragraph.font.color.rgb = color
        paragraph.font.bold = bold


def _fill_placeholders(slide, title_text: str, body_text: str):
    """
    Fill placeholders in slide with provided title and body. If not present, create textboxes.
    """
    from pptx.enum.shapes import PP_PLACEHOLDER
    title_filled = False
    body_filled = False

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        try:
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type
                if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) and not title_filled:
                    shape.text_frame.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = title_text
                    _set_paragraph_font(p, TITLE_FONT_NAME, TITLE_PT, TITLE_COLOR, bold=True)
                    title_filled = True
                elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT) and not body_filled:
                    tf = shape.text_frame
                    tf.clear()
                    paragraphs = [p.strip() for p in body_text.split("\n\n") if p.strip()]
                    for i, para in enumerate(paragraphs):
                        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                        if para.startswith("- ") or para.startswith("* "):
                            p.text = para[2:].strip()
                            p.level = 0
                        else:
                            p.text = para
                        _set_paragraph_font(p, BODY_FONT_NAME, BODY_PT, BODY_COLOR, bold=False)
                    body_filled = True
        except Exception:
            pass

    if not title_filled:
        # Add textbox for title
        left, top, width, height = Inches(0.5), Inches(0.3), Inches(9.0), Inches(1.0)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = title_text
        _set_paragraph_font(p, TITLE_FONT_NAME, TITLE_PT, TITLE_COLOR, bold=True)

    if not body_filled:
        left, top, width, height = Inches(0.5), Inches(1.4), Inches(9.0), Inches(5.0)
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        tf.clear()
        paragraphs = [p.strip() for p in body_text.split("\n\n") if p.strip()]
        for i, para in enumerate(paragraphs):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            if para.startswith("- ") or para.startswith("* "):
                p.text = para[2:].strip()
                p.level = 0
            else:
                p.text = para
            _set_paragraph_font(p, BODY_FONT_NAME, BODY_PT, BODY_COLOR, bold=False)


def fill_template_with_pages(
    template_bytes: Optional[bytes],
    pages: List[Dict[str, str]],
    title_font: str = TITLE_FONT_NAME,
    body_font: str = BODY_FONT_NAME,
    title_font_pt: Optional[float] = None,
    body_font_pt: Optional[float] = None,
) -> bytes:
    """
    Create a new Presentation by filling the provided pages into the template (if given) or into a simple blank layout.
    """
    if template_bytes:
        out_prs = Presentation(io.BytesIO(template_bytes))
        # remove existing slides while preserving masters/layouts
        sldIdLst = out_prs.slides._sldIdLst  # pylint: disable=protected-access
        for sldId in list(sldIdLst):
            sldIdLst.remove(sldId)
        layout_idx = find_layout_index_with_title_and_body(out_prs)
        for p in pages:
            slide = out_prs.slides.add_slide(out_prs.slide_layouts[layout_idx])
            _fill_placeholders(slide, p["title"], p["body"])
    else:
        out_prs = Presentation()
        # Try to match page size to default Presentation
        for p in pages:
            blank_layout = out_prs.slide_layouts[6] if len(out_prs.slide_layouts) > 6 else out_prs.slide_layouts[0]
            slide = out_prs.slides.add_slide(blank_layout)
            _fill_placeholders(slide, p["title"], p["body"])

    out_stream = io.BytesIO()
    out_prs.save(out_stream)
    return out_stream.getvalue()
