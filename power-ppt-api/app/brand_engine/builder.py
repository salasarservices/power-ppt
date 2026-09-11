"""
Public entry point: turn a SlidePlan into a valid, on-brand PPTX against the
bundled template. The flow planner packs each page's content across as many
slides as it needs (overflow continuation); the integrity guards run before
returning bytes.
"""

import io

from .content import render_placements
from .errors import BrandEngineError
from .flow import flow_pages
from .heading import render_heading
from .integrity import verify_output
from .template import (
    clone_content_slide,
    fix_shape_ids,
    load_template,
    remove_slide,
    remove_slide_number_fields,
)


def build_deck(plan, template_path) -> bytes:
    """
    Build the standardised deck. `plan` is a SlidePlan (or any object exposing the
    same attributes). Raises BrandEngineError if the result fails the guards.
    """
    prs = load_template(template_path)
    if not prs.slides:
        raise BrandEngineError("Brand template has no slides.")

    n_template = len(prs.slides)
    specs = flow_pages(plan.pages)
    if not specs:
        raise BrandEngineError("SlidePlan produced no pages.")

    for spec in specs:
        slide = clone_content_slide(prs, source_idx=0)
        remove_slide_number_fields(slide)
        render_heading(slide, spec["title"])
        render_placements(slide, spec["placements"])
        fix_shape_ids(slide)

    # Remove the original template slides (still at the front).
    for _ in range(n_template):
        remove_slide(prs, 0)

    buf = io.BytesIO()
    prs.save(buf)
    data = buf.getvalue()

    verify_output(data, specs, template_path)
    return data
