"""
Content flow planner (A2). Packs each page's blocks (body, tables, images) down
the content zone and spills the remainder onto continuation slides that repeat the
title with a "(CONTD...)" suffix. Tables split by rows (header repeated); long
paragraphs split by words; images are atomic (scaled to one zone if oversized).

Works in estimated inches — the same estimators the renderer uses — so a block the
planner puts on a slide is a block the renderer can fit.
"""

import base64

from ..schemas import Table
from . import content as c
from . import geometry as g

_ZONE_H = g.BODY_HEIGHT           # usable content-zone height (inches)
_GAP = g.CONTENT_GAP
_EPS = 1e-6


def _decode_images(page) -> list[bytes]:
    out: list[bytes] = []
    for img in getattr(page, "images", None) or []:
        try:
            out.append(base64.b64decode(img.data))
        except Exception:
            pass
    return out


def _page_blocks(page) -> list[tuple]:
    blocks: list[tuple] = []
    body = (page.body or "").strip()
    if body:
        blocks.append(("body", body))
    for t in page.tables or []:
        if list(t.rows or []):
            blocks.append(("table", t))
    for img in _decode_images(page):
        blocks.append(("image", img))
    return blocks


def _split_paragraph_by_words(para: str, height: float) -> tuple[str, str]:
    words = para.split()
    acc: list[str] = []
    for k, w in enumerate(words):
        if c.estimate_body_height_in(" ".join(acc + [w])) <= height or not acc:
            acc.append(w)
        else:
            return " ".join(acc), " ".join(words[k:])
    return para, ""


def _split_body(text: str, height: float) -> tuple[str, str | None]:
    """Return (fits_in_height, remainder|None), splitting by paragraph then word."""
    paras = text.split("\n\n")
    acc: list[str] = []
    i = 0
    while i < len(paras):
        if c.estimate_body_height_in("\n\n".join(acc + [paras[i]])) <= height:
            acc.append(paras[i])
            i += 1
        else:
            break
    if i == len(paras):
        return text, None
    if not acc:  # first paragraph alone overflows -> split it by words
        head, tail = _split_paragraph_by_words(paras[i], height)
        rest = ([tail] if tail else []) + paras[i + 1:]
        return head, ("\n\n".join(p for p in rest if p) or None)
    return "\n\n".join(acc), "\n\n".join(paras[i:])


def _split_table(table, height: float):
    """Return (placed|None, remainder|None). Header (rows[0]) repeats on each part."""
    rows = list(table.rows or [])
    header = rows[0] if rows else list(table.header or [])
    data = rows[1:]
    max_rows = int(height / g.TABLE_ROW_H + _EPS)
    if max_rows >= len(rows):
        return table, None
    usable = max_rows - 1                      # reserve one row for the header
    if usable < 1:
        return None, table                     # not even header+1 row fits here
    placed = Table(header=header, rows=[header] + data[:usable])
    rest = data[usable:]
    remainder = Table(header=header, rows=[header] + rest) if rest else None
    return placed, remainder


def _pack(blocks: list[tuple]) -> list[list[tuple]]:
    slides: list[list[tuple]] = []
    cur: list[tuple] = []
    remaining = _ZONE_H

    def new_slide():
        nonlocal cur, remaining
        slides.append(cur)
        cur = []
        remaining = _ZONE_H

    for kind, payload in blocks:
        while payload is not None:
            fresh = remaining >= _ZONE_H - _EPS

            if kind == "body":
                placed, payload = _split_body(payload, remaining)
                if not placed:
                    if fresh:                  # nothing fits even on a blank slide
                        cur.append((kind, payload)); payload = None
                    else:
                        new_slide(); continue
                else:
                    cur.append((kind, placed))
                    remaining -= c.estimate_body_height_in(placed) + _GAP
                    if payload is not None:
                        new_slide()

            elif kind == "table":
                placed, payload = _split_table(payload, remaining)
                if placed is None:
                    if fresh:                  # zone too small even for header+1: force
                        cur.append((kind, payload)); payload = None
                    else:
                        new_slide(); continue
                else:
                    cur.append((kind, placed))
                    remaining -= len(list(placed.rows)) * g.TABLE_ROW_H + _GAP
                    if payload is not None:
                        new_slide()

            else:  # image (atomic)
                h = c.estimate_image_height_in(payload)
                if h <= remaining or fresh:    # fits, or blank slide (renderer clamps)
                    cur.append((kind, payload))
                    remaining -= min(h, _ZONE_H) + _GAP
                    payload = None
                else:
                    new_slide()

    slides.append(cur)                          # flush (may be empty -> title-only slide)
    return slides


def flow_pages(pages) -> list[dict]:
    """Flatten pages into physical slide specs: {title, blocks}. Continuation
    slides carry the same title + CONTINUATION_SUFFIX."""
    out: list[dict] = []
    for page in pages:
        title = page.title or ""
        for idx, blocks in enumerate(_pack(_page_blocks(page))):
            if idx == 0:
                t = title
            else:
                t = (title + g.CONTINUATION_SUFFIX) if title else ""
            out.append({"title": t, "blocks": blocks})
    return out
