import io

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

# Decorative/standing text to exclude from body content (brand tagline).
EXCLUDED_PHRASES = [
    "you manage your enterprise",
    "we will manage your insurable risks",
]

# Title placeholder types (their text is the slide title, not body).
_TITLE_PH = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE}

# Placeholders that are chrome, never content (would otherwise pollute the body).
_NONCONTENT_PH = {
    PP_PLACEHOLDER.SLIDE_NUMBER,
    PP_PLACEHOLDER.FOOTER,
    PP_PLACEHOLDER.HEADER,
    PP_PLACEHOLDER.DATE,
}


def should_exclude_text(text: str) -> bool:
    t = text.strip().lower()
    return any(p in t for p in EXCLUDED_PHRASES)


def _ph_type(shape):
    if not getattr(shape, "is_placeholder", False):
        return None
    try:
        return shape.placeholder_format.type
    except Exception:
        return None


def _iter_shapes(shapes):
    """Flatten a shape tree, recursing into groups so nested text/tables/images
    are not missed (bug #5 fix)."""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(sh.shapes)
        else:
            yield sh


def extract_text_shapes(pptx_bytes):
    """
    Extract title, body text, tables and images from every slide. Recurses groups,
    reads any shape's text (text boxes, placeholders, freeform/autoshapes with a
    text frame), and ignores chrome placeholders (slide number/footer/date/header).
    Title comes only from a real title placeholder — no "Slide N" fallback.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slides_meta = []

    for slide_idx, slide in enumerate(prs.slides):
        title_text = ""
        body_parts = []
        image_shapes = []
        shapes = []

        for shape in _iter_shapes(slide.shapes):
            shapes.append(shape)
            pht = _ph_type(shape)

            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    if pht in _TITLE_PH:
                        if not title_text:
                            title_text = text
                    elif pht in _NONCONTENT_PH:
                        pass  # slide number / footer / date — never body
                    elif not should_exclude_text(text):
                        body_parts.append(text)

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    img = shape.image
                    image_shapes.append(
                        {"image_bytes": img.blob, "content_type": img.content_type}
                    )
                except Exception as e:
                    print(f"Failed to extract image from slide {slide_idx}: {e}")

        slides_meta.append({
            "slide_index": slide_idx,
            "title_text": title_text,           # "" when the slide has no title
            "body_text": "\n\n".join(body_parts),
            "image_shapes": image_shapes,
            "shapes": shapes,                   # flattened; tables read from here
        })

    return slides_meta
