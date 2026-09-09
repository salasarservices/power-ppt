"""
Public entry point: turn a SlidePlan into a valid, on-brand PPTX against the
bundled template. Handles body-overflow pagination and runs the integrity guards
before returning bytes.
"""

import io

from ..paginator import split_by_paragraphs
from . import geometry as g
from .content import render_body, render_table
from .errors import BrandEngineError
from .heading import render_heading
from .integrity import verify_output
from .template import (
    clone_content_slide,
    fix_shape_ids,
    load_template,
    remove_slide,
    remove_slide_number_fields,
)


def _expand_pages(plan) -> list[dict]:
    """
    Flatten a SlidePlan into rendered pages, paginating long bodies. A page with
    tables is kept whole (tables are not paginated in v1).
    """
    out: list[dict] = []
    for page in plan.pages:
        title = page.title or ""
        tables = list(page.tables or [])
        body = page.body or ""

        if tables:
            out.append({"title": title, "body": body, "tables": tables})
            continue

        chunks = (
            split_by_paragraphs(body, g.DEFAULT_CHARS_PER_PAGE)
            if body.strip()
            else [""]
        )
        for i, chunk in enumerate(chunks):
            if i == 0:
                page_title = title
            else:
                page_title = (title + g.CONTINUATION_SUFFIX) if title else ""
            out.append({"title": page_title, "body": chunk, "tables": []})

    return out


def build_deck(plan, template_path) -> bytes:
    """
    Build the standardised deck. `plan` is a SlidePlan (or any object exposing the
    same attributes). Raises BrandEngineError if the result fails the guards.
    """
    prs = load_template(template_path)
    if not prs.slides:
        raise BrandEngineError("Brand template has no slides.")

    n_template = len(prs.slides)
    rendered = _expand_pages(plan)
    if not rendered:
        raise BrandEngineError("SlidePlan produced no pages.")

    for page in rendered:
        slide = clone_content_slide(prs, source_idx=0)
        remove_slide_number_fields(slide)
        render_heading(slide, page["title"])
        if page["tables"]:
            for tbl in page["tables"]:
                render_table(slide, tbl)
        else:
            render_body(slide, page["body"])
        fix_shape_ids(slide)

    # Remove the original template slides (still at the front).
    for _ in range(n_template):
        remove_slide(prs, 0)

    buf = io.BytesIO()
    prs.save(buf)
    data = buf.getvalue()

    verify_output(data, rendered, template_path)
    return data
