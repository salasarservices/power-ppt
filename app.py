# app.py
import io
from typing import List, Tuple, Optional

import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt

# Styling constants
TITLE_FONT_NAME = "Poppins"
BODY_FONT_NAME = "Poppins"
TITLE_PX = 20
BODY_PX = 12
# Convert px to pt (approx): 1 px = 0.75 pt
TITLE_PT = Pt(TITLE_PX * 0.75)
BODY_PT = Pt(BODY_PX * 0.75)
TITLE_COLOR = RGBColor(0x2d, 0x44, 0x8d)  # #2d448d
BODY_COLOR = RGBColor(0x00, 0x00, 0x00)   # #000000

# Default layout positions if template doesn't provide placeholders
TITLE_LEFT = Inches(0.5)
TITLE_TOP = Inches(0.3)
TITLE_WIDTH = Inches(9.0)
TITLE_HEIGHT = Inches(1.0)

BODY_LEFT = Inches(0.5)
BODY_TOP = Inches(1.4)
BODY_WIDTH = Inches(9.0)
BODY_HEIGHT = Inches(5.0)

# Pagination heuristic
CHARS_PER_PAGE = 1100


def extract_title_and_body(slide) -> Tuple[str, str]:
    """
    Extract title and body text heuristically:
    - Prefer TITLE placeholder for title.
    - For body, gather text from other text shapes excluding footers/date/slide number and the title shape.
    """
    title = ""
    # Try to find placeholder title
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        try:
            if shape.is_placeholder and shape.placeholder_format.type in (
                PP_PLACEHOLDER.TITLE,
                PP_PLACEHOLDER.CENTER_TITLE,
            ):
                if shape.text.strip():
                    title = shape.text.strip()
                    break
        except Exception:
            pass

    text_shapes = []
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        try:
            if shape.is_placeholder and shape.placeholder_format.type in (
                PP_PLACEHOLDER.SLIDE_NUMBER,
                PP_PLACEHOLDER.FOOTER,
                PP_PLACEHOLDER.DATE,
            ):
                continue
        except Exception:
            pass
        text = shape.text.strip()
        if not text:
            continue
        # Skip the title shape if we've captured title
        if title and text == title:
            continue
        text_shapes.append(text)

    # If no title found, use first shape's first non-empty line
    if not title:
        if text_shapes:
            first = text_shapes.pop(0)
            lines = [l for l in first.splitlines() if l.strip()]
            if lines:
                title = lines[0].strip()
                rest = "\n".join(lines[1:]).strip()
                if rest:
                    text_shapes.insert(0, rest)
        else:
            title = "Untitled"

    body = "\n\n".join(text_shapes).strip()
    return title, body


def split_text_into_pages(text: str, chars_per_page: int = CHARS_PER_PAGE) -> List[str]:
    """
    Split the body text into pages preserving paragraph boundaries.
    """
    if not text:
        return [""]
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    pages = []
    current = ""
    for p in paragraphs:
        if current and (len(current) + len(p) + 2) > chars_per_page:
            pages.append(current.strip())
            current = p
        else:
            if current:
                current += "\n\n" + p
            else:
                current = p
    if current.strip():
        pages.append(current.strip())
    if not pages:
        pages = [""]
    return pages


def clear_all_slides(prs: Presentation):
    """
    Remove all slides from a Presentation (keeps masters and layouts).
    """
    sldIdLst = prs.slides._sldIdLst  # pylint: disable=protected-access
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)


def find_layout_index_with_title_and_body(prs: Presentation) -> int:
    """
    Return index of a layout that has both title and body/content placeholders.
    Fallback to 0.
    """
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
    # fallback: try to pick layout that has a title
    for idx, layout in enumerate(prs.slide_layouts):
        for ph in layout.placeholders:
            try:
                if ph.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                    return idx
            except Exception:
                pass
    return 0


def set_paragraph_font(paragraph, name: str, size: Pt, color: RGBColor, bold: bool = False):
    for run in paragraph.runs:
        run.font.name = name
        run.font.size = size
        run.font.color.rgb = color
        run.font.bold = bold
    # If paragraph has no runs (rare), set on paragraph level
    if not paragraph.runs:
        paragraph.font.name = name
        paragraph.font.size = size
        paragraph.font.color.rgb = color
        paragraph.font.bold = bold


def fill_slide_placeholders_with_title_and_body(slide, title_text: str, body_text: str) -> None:
    """
    Fill title and body placeholders on the slide if they exist.
    If no appropriate placeholders exist, create textboxes as fallback.
    """
    title_filled = False
    body_filled = False

    # Try to fill placeholders
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
                    set_paragraph_font(p, TITLE_FONT_NAME, TITLE_PT, TITLE_COLOR, bold=True)
                    title_filled = True
                elif ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT) and not body_filled:
                    # Clear and populate paragraphs
                    tf = shape.text_frame
                    tf.clear()
                    paragraphs = [p.strip() for p in body_text.split("\n\n") if p.strip()]
                    for i, para in enumerate(paragraphs):
                        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                        # preserve simple bullet markers like "- " or "* "
                        if para.startswith("- ") or para.startswith("* "):
                            p.text = para[2:].strip()
                            p.level = 0
                        else:
                            p.text = para
                        set_paragraph_font(p, BODY_FONT_NAME, BODY_PT, BODY_COLOR, bold=False)
                    body_filled = True
        except Exception:
            # If any placeholder access error occurs, skip this shape
            pass

    # Fallbacks: if placeholders not filled, create textboxes at default positions
    if not title_filled:
        title_box = slide.shapes.add_textbox(TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT)
        tf_title = title_box.text_frame
        tf_title.clear()
        p = tf_title.paragraphs[0]
        p.text = title_text
        set_paragraph_font(p, TITLE_FONT_NAME, TITLE_PT, TITLE_COLOR, bold=True)

    if not body_filled:
        body_box = slide.shapes.add_textbox(BODY_LEFT, BODY_TOP, BODY_WIDTH, BODY_HEIGHT)
        tf_body = body_box.text_frame
        tf_body.clear()
        paragraphs = [p.strip() for p in body_text.split("\n\n") if p.strip()]
        for i, para in enumerate(paragraphs):
            p = tf_body.add_paragraph() if i > 0 else tf_body.paragraphs[0]
            if para.startswith("- ") or para.startswith("* "):
                p.text = para[2:].strip()
                p.level = 0
            else:
                p.text = para
            set_paragraph_font(p, BODY_FONT_NAME, BODY_PT, BODY_COLOR, bold=False)


def process_pptx_bytes(input_bytes: bytes, template_bytes: Optional[bytes] = None) -> bytes:
    """
    Process input presentation and produce a standardized presentation.
    If template_bytes is provided, use that template's slide layouts/backgrounds.
    Otherwise, use a default blank layout and textboxes.
    """
    in_prs = Presentation(io.BytesIO(input_bytes))

    if template_bytes:
        out_prs = Presentation(io.BytesIO(template_bytes))
        # Clear any existing slides from template so we can add fresh ones while preserving masters/layouts
        clear_all_slides(out_prs)
        layout_idx = find_layout_index_with_title_and_body(out_prs)
        use_template = True
    else:
        out_prs = Presentation()
        layout_idx = None
        use_template = False
        # Ensure output size matches input for consistent aspect ratio
        out_prs.slide_width = in_prs.slide_width
        out_prs.slide_height = in_prs.slide_height

    for slide in in_prs.slides:
        title, body = extract_title_and_body(slide)
        if not body:
            pages = [""]
        else:
            pages = split_text_into_pages(body)
        for i, page_text in enumerate(pages):
            page_title = title if i == 0 else f"{title} (cont.)"
            if use_template:
                new_slide = out_prs.slides.add_slide(out_prs.slide_layouts[layout_idx])
                fill_slide_placeholders_with_title_and_body(new_slide, page_title, page_text)
            else:
                # create blank slide and add textboxes
                blank_layout = out_prs.slide_layouts[6] if len(out_prs.slide_layouts) > 6 else out_prs.slide_layouts[0]
                new_slide = out_prs.slides.add_slide(blank_layout)
                fill_slide_placeholders_with_title_and_body(new_slide, page_title, page_text)

    out_stream = io.BytesIO()
    out_prs.save(out_stream)
    return out_stream.getvalue()


# Streamlit UI
st.set_page_config(page_title="PPT Normalizer — Template Upload", layout="centered")
st.title("PPT Normalizer — Use a Runtime Template")

st.markdown(
    """
Upload an input PowerPoint (.pptx) and optionally upload a standardized template (.pptx).
The app will extract each slide's Title and main Body content from the input and place them into the template's slide layout (preserving background and placeholders).
- Title: Poppins (Semibold), 20 px, color #2d448d
- Body: Poppins (Regular), 12 px, color #000000

If the template is not provided, a simple standardized layout is used instead.

Notes:
- python-pptx does not embed fonts. To see Poppins exactly, ensure Poppins is installed on the machine that opens the final PPTX.
- The app extracts only textual content (title + body). Complex objects (charts, images, tables) are not transferred.
"""
)

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Upload the input .pptx file", type=["pptx"])
with col2:
    template_file = st.file_uploader("Upload standardized template (.pptx) — optional", type=["pptx"])

if uploaded_file is not None:
    input_bytes = uploaded_file.read()
    template_bytes = template_file.read() if template_file is not None else None
    st.info("Processing uploaded PPTX...")
    try:
        output_bytes = process_pptx_bytes(input_bytes, template_bytes=template_bytes)
        st.success("Processed successfully.")
        st.download_button(
            label="Download standardized PPTX",
            data=output_bytes,
            file_name=f"standardized_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        st.markdown(
            """
Notes:
- The app sets fonts to 'Poppins' on the text elements. To see exact typography, ensure Poppins is installed on the target machine.
- If you want the app to use a specific slide index or layout from your template, tell me and I can add an input option to pick which template layout to use.
"""
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("Upload an input .pptx file to get started. You may also upload a template (.pptx) to control styling/background.")
