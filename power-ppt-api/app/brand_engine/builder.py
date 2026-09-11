"""
Public entry points:
  flow_deck(pages)          SlidePlan pages -> positioned Deck (Placement JSON)
  render_deck(deck, tpl)    Deck -> valid, on-brand PPTX bytes (WYSIWYG)
  build_deck(plan, tpl)     convenience: flow_deck then render_deck

The Deck (Placement JSON) is the editor's round-trip contract: /layout returns it,
the canvas renders and edits it, /render turns the exact placements into the .pptx.
"""

import base64
import io

from ..schemas import Deck, Image, Placement, Slide
from .content import estimate_body_height_in, render_placements
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


def _specs_to_deck(specs) -> Deck:
    """Internal flow specs (raw payloads) -> serializable Deck (base64 images)."""
    slides: list[Slide] = []
    for spec in specs:
        pls: list[Placement] = []
        for p in spec["placements"]:
            k = p["kind"]
            if k == "body":
                pls.append(Placement(kind="body", text=p["payload"], left=p["left"],
                                     top=p["top"], width=p["width"], height=p["height"]))
            elif k == "table":
                pls.append(Placement(kind="table", table=p["payload"], left=p["left"],
                                     top=p["top"], width=p["width"]))
            elif k == "image":
                pls.append(Placement(
                    kind="image",
                    image=Image(data=base64.b64encode(p["payload"]).decode("ascii")),
                    left=p["left"], top=p["top"], width=p["width"],
                ))
        slides.append(Slide(title=spec["title"], placements=pls))
    return Deck(slides=slides)


def _deck_to_specs(deck: Deck) -> list[dict]:
    """Deck -> internal render specs (decoded image bytes, body height filled in)."""
    specs: list[dict] = []
    for sl in deck.slides:
        pls: list[dict] = []
        for p in sl.placements:
            if p.kind == "body":
                text = p.text or ""
                pls.append({"kind": "body", "payload": text, "left": p.left, "top": p.top,
                            "width": p.width,
                            "height": p.height or estimate_body_height_in(text, p.width)})
            elif p.kind == "table" and p.table is not None:
                pls.append({"kind": "table", "payload": p.table, "left": p.left,
                            "top": p.top, "width": p.width})
            elif p.kind == "image" and p.image is not None:
                try:
                    data = base64.b64decode(p.image.data)
                except Exception:
                    continue
                pls.append({"kind": "image", "payload": data, "left": p.left,
                            "top": p.top, "width": p.width})
        specs.append({"title": sl.title, "placements": pls})
    return specs


def flow_deck(pages) -> Deck:
    """Flow content pages into a positioned Deck (Placement JSON)."""
    return _specs_to_deck(flow_pages(pages))


def render_deck(deck: Deck, template_path) -> bytes:
    """Render a Deck's placements into a valid, on-brand PPTX. Raises
    BrandEngineError if the result fails the integrity guards."""
    prs = load_template(template_path)
    if not prs.slides:
        raise BrandEngineError("Brand template has no slides.")
    if not deck.slides:
        raise BrandEngineError("Deck has no slides.")

    n_template = len(prs.slides)
    specs = _deck_to_specs(deck)

    for spec in specs:
        slide = clone_content_slide(prs, source_idx=0)
        remove_slide_number_fields(slide)
        render_heading(slide, spec["title"])
        render_placements(slide, spec["placements"])
        fix_shape_ids(slide)

    for _ in range(n_template):
        remove_slide(prs, 0)

    buf = io.BytesIO()
    prs.save(buf)
    data = buf.getvalue()

    verify_output(data, specs, template_path)
    return data


def build_deck(plan, template_path) -> bytes:
    """Convenience: flow a SlidePlan and render it in one call."""
    deck = flow_deck(plan.pages)
    if not deck.slides:
        raise BrandEngineError("SlidePlan produced no pages.")
    return render_deck(deck, template_path)
