# paginator.py
import io
import textwrap
from typing import List, Optional

from PIL import ImageFont, Image, ImageDraw


def split_by_paragraphs(text: str, chars_per_page: int = 1100) -> List[str]:
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


def _measure_text_lines(text: str, font: ImageFont.FreeTypeFont, max_width_px: int) -> List[str]:
    # Wrap into lines using textwrap approximations, then refine using width measure
    wrapper = textwrap.TextWrapper(width=200)
    paragraphs = text.split("\n\n")
    lines = []
    for para in paragraphs:
        if not para.strip():
            continue
        rough = wrapper.wrap(para)
        for r in rough:
            # refine wrap by measuring
            words = r.split(" ")
            cur = ""
            for w in words:
                if cur == "":
                    cur = w
                else:
                    test = cur + " " + w
                    w_px = font.getsize(test)[0]
                    if w_px <= max_width_px:
                        cur = test
                    else:
                        lines.append(cur)
                        cur = w
            if cur:
                lines.append(cur)
        # preserve paragraph break as empty line
        lines.append("")
    return lines


def split_by_font_metrics(
    text: str,
    font_bytes: bytes,
    box_width_px: Optional[int],
    box_height_px: Optional[int],
    font_size_px: int = 12,
    dpi: int = 300,
) -> List[str]:
    """
    Precise pagination using font metrics. If box_width_px/box_height_px are None, fallback to default heuristic.
    """
    if not text:
        return [""]
    if box_width_px is None or box_height_px is None:
        # Can't do precise pagination without a target box; fallback
        return split_by_paragraphs(text)

    # Load font from bytes
    font = ImageFont.truetype(io.BytesIO(font_bytes), size=font_size_px)
    # Produce lines
    lines = _measure_text_lines(text, font, box_width_px)
    # Now fill pages by height
    # Estimate line height
    ascent, descent = font.getmetrics()
    line_height = ascent + descent + 2  # small padding
    pages = []
    current = []
    current_h = 0
    for line in lines:
        # if it's paragraph break
        if line == "":
            # account for a blank-line height
            needed = line_height
        else:
            needed = line_height
        if current_h + needed > box_height_px and current:
            pages.append("\n".join(current).strip())
            current = []
            current_h = 0
        if line == "":
            current.append("")  # paragraph break
            current_h += needed
        else:
            current.append(line)
            current_h += needed
    if current:
        pages.append("\n".join(current).strip())
    if not pages:
        pages = [""]
    return pages
