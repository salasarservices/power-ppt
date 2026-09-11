"""
Package-2 integrity guards. A corrupt or incomplete deck must never be returned:
build_deck runs verify_output on its own output and raises BrandEngineError on any
failure.
"""

import hashlib
import io

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .errors import BrandEngineError


def _iter_pictures(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield sh
        elif sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_pictures(sh.shapes)


def _slide_text(shapes) -> str:
    parts = []
    for sh in shapes:
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if sh.has_table:
            for row in sh.table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            parts.append(_slide_text(sh.shapes))
    return "\n".join(parts)


def _cnvpr_ids(slide):
    return [
        elem.get("id")
        for elem in slide.shapes._spTree.iter()
        if elem.tag.endswith("}cNvPr")
    ]


def _template_image_hashes(template_path):
    prs = Presentation(template_path)
    return {
        hashlib.md5(pic.image.blob).hexdigest()
        for pic in _iter_pictures(prs.slides[0].shapes)
    }


def _spec_body(spec) -> str:
    """Reconstruct the body text a slide spec should contain from its placements."""
    return "\n\n".join(
        p["payload"] for p in spec.get("placements", []) if p.get("kind") == "body"
    )


def verify_output(pptx_bytes: bytes, rendered_pages: list, template_path) -> None:
    """Raise BrandEngineError if the output deck is invalid or off-brand."""
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
    except Exception as e:  # invalid XML / not a PPTX
        raise BrandEngineError(f"Output is not a valid PPTX: {e}") from e

    slides = list(prs.slides)
    if len(slides) != len(rendered_pages):
        raise BrandEngineError(
            f"Slide count {len(slides)} != expected {len(rendered_pages)}"
        )

    tpl_hashes = _template_image_hashes(template_path)

    for i, (slide, page) in enumerate(zip(slides, rendered_pages), start=1):
        ids = _cnvpr_ids(slide)
        if len(ids) != len(set(ids)):
            raise BrandEngineError(f"Slide {i} has duplicate shape IDs")

        pic_hashes = {
            hashlib.md5(p.image.blob).hexdigest()
            for p in _iter_pictures(slide.shapes)
        }
        if tpl_hashes - pic_hashes:
            raise BrandEngineError(f"Slide {i} is missing a brand image (logo/background)")

        text = _slide_text(slide.shapes).upper()

        title = (page.get("title") or "").strip().upper()
        if title:
            first_word = title.split(" ", 1)[0]
            if first_word and first_word not in text:
                raise BrandEngineError(f"Slide {i} is missing its title text")
        else:
            body = _spec_body(page).strip()
            if body:
                token = body.split()[0].upper()
                if token and token not in text:
                    raise BrandEngineError(f"Slide {i} is missing its body content")
