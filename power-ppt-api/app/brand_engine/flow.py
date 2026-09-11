"""
Content flow planner (A2 + 2D layout). For each page it:
  1. Builds layout units — body is a full-width block; tables/images are packed
     into rows, placing two side-by-side when both fit at half-width (e.g. a table
     beside an image), else one full-width.
  2. Packs those units down the content zone, spilling the remainder onto
     continuation slides titled "<title> (CONTD...)". Full-width tables split by
     rows (header repeated); body splits by paragraph/word; images are atomic.

Output: per-slide {title, placements}, where each placement is a positioned dict
{kind, payload, left, top, width, (height for body)} the renderer draws directly.
Works in estimated inches — the same estimators the renderer uses.
"""

import base64

from ..schemas import Table
from . import content as c
from . import geometry as g

_ZONE_H = g.BODY_HEIGHT
_GAP = g.CONTENT_GAP
_EPS = 1e-6


# ── input decode / block build ───────────────────────────────────────────────
def _decode_images(page) -> list[bytes]:
    out: list[bytes] = []
    for img in getattr(page, "images", None) or []:
        try:
            out.append(base64.b64decode(img.data))
        except Exception:
            pass
    return out


def _half_friendly(kind, payload) -> bool:
    if kind == "image":
        return True
    if kind == "table":
        ncols = max((len(r) for r in (payload.rows or [])), default=0)
        return 0 < ncols <= g.HALF_TABLE_MAX_COLS
    return False


def _cell_height(kind, payload, width) -> float:
    if kind == "table":
        return c.estimate_table_height_in(payload)
    if kind == "image":
        return c.estimate_image_height_in(payload, width)
    return 0.0


def _build_units(page) -> list[tuple]:
    """Return ordered units: ("body", text) | ("row", [(kind,payload,left,width), ...])."""
    units: list[tuple] = []
    body = (page.body or "").strip()
    if body:
        units.append(("body", body))

    packables = [("table", t) for t in (page.tables or []) if list(t.rows or [])]
    packables += [("image", img) for img in _decode_images(page)]

    i = 0
    while i < len(packables):
        a = packables[i]
        b = packables[i + 1] if i + 1 < len(packables) else None
        if (
            b
            and _half_friendly(*a)
            and _half_friendly(*b)
            and max(_cell_height(*a, g.HALF_WIDTH), _cell_height(*b, g.HALF_WIDTH)) <= _ZONE_H
        ):
            units.append(("row", [
                (a[0], a[1], g.BODY_LEFT, g.HALF_WIDTH),
                (b[0], b[1], g.BODY_LEFT + g.HALF_WIDTH + g.COL_GUTTER, g.HALF_WIDTH),
            ]))
            i += 2
        else:
            units.append(("row", [(a[0], a[1], g.BODY_LEFT, g.BODY_WIDTH)]))
            i += 1
    return units


# ── splitting helpers ────────────────────────────────────────────────────────
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
    if not acc:
        head, tail = _split_paragraph_by_words(paras[i], height)
        rest = ([tail] if tail else []) + paras[i + 1:]
        return head, ("\n\n".join(p for p in rest if p) or None)
    return "\n\n".join(acc), "\n\n".join(paras[i:])


def _split_table(table, height: float):
    rows = list(table.rows or [])
    header = rows[0] if rows else list(table.header or [])
    data = rows[1:]
    max_rows = int(height / g.TABLE_ROW_H + _EPS)
    if max_rows >= len(rows):
        return table, None
    usable = max_rows - 1
    if usable < 1:
        return None, table
    placed = Table(header=header, rows=[header] + data[:usable])
    rest = data[usable:]
    return placed, (Table(header=header, rows=[header] + rest) if rest else None)


# ── packer ───────────────────────────────────────────────────────────────────
class _Packer:
    def __init__(self):
        self.slides: list[list[dict]] = []
        self.cur: list[dict] = []
        self.y = g.BODY_TOP
        self.remaining = _ZONE_H

    def _new_slide(self):
        self.slides.append(self.cur)
        self.cur = []
        self.y = g.BODY_TOP
        self.remaining = _ZONE_H

    def _fresh(self) -> bool:
        return self.remaining >= _ZONE_H - _EPS

    def _advance(self, h: float):
        self.y += h + _GAP
        self.remaining -= h + _GAP

    def add_body(self, text: str):
        payload: str | None = text
        while payload is not None:
            placed, payload = _split_body(payload, self.remaining)
            if not placed:
                if self._fresh():        # can't fit even blank slide -> force
                    placed, payload = payload, None
                else:
                    self._new_slide()
                    continue
            h = c.estimate_body_height_in(placed)
            self.cur.append({"kind": "body", "payload": placed, "left": g.BODY_LEFT,
                             "top": self.y, "width": g.BODY_WIDTH, "height": h})
            self._advance(h)
            if payload is not None:
                self._new_slide()

    def add_table_fullwidth(self, table):
        payload = table
        while payload is not None:
            placed, payload = _split_table(payload, self.remaining)
            if placed is None:
                if self._fresh():
                    placed, payload = payload, None
                else:
                    self._new_slide()
                    continue
            h = c.estimate_table_height_in(placed)
            self.cur.append({"kind": "table", "payload": placed, "left": g.BODY_LEFT,
                             "top": self.y, "width": g.BODY_WIDTH})
            self._advance(h)
            if payload is not None:
                self._new_slide()

    def add_image_fullwidth(self, img):
        h = c.estimate_image_height_in(img, g.BODY_WIDTH)
        if h > self.remaining and not self._fresh():
            self._new_slide()
        self.cur.append({"kind": "image", "payload": img, "left": g.BODY_LEFT,
                         "top": self.y, "width": g.BODY_WIDTH})
        self._advance(min(h, _ZONE_H))       # renderer clamps oversized images to the zone

    def add_row(self, cells):
        rh = max(_cell_height(k, p, w) for (k, p, _l, w) in cells)
        if rh > self.remaining and not self._fresh():
            self._new_slide()
        for (k, p, l, w) in cells:
            self.cur.append({"kind": k, "payload": p, "left": l, "top": self.y, "width": w})
        self._advance(rh)

    def finish(self) -> list[list[dict]]:
        self.slides.append(self.cur)         # flush (may be empty -> title-only slide)
        return self.slides


def _pack(units) -> list[list[dict]]:
    pk = _Packer()
    for kind, data in units:
        if kind == "body":
            pk.add_body(data)
        else:  # row
            cells = data
            if len(cells) == 1 and cells[0][0] == "table":
                pk.add_table_fullwidth(cells[0][1])
            elif len(cells) == 1 and cells[0][0] == "image":
                pk.add_image_fullwidth(cells[0][1])
            else:
                pk.add_row(cells)
    return pk.finish()


def flow_pages(pages) -> list[dict]:
    out: list[dict] = []
    for page in pages:
        title = page.title or ""
        for idx, placements in enumerate(_pack(_build_units(page))):
            t = title if idx == 0 else ((title + g.CONTINUATION_SUFFIX) if title else "")
            out.append({"title": t, "placements": placements})
    return out
