# utils.py
from typing import List, Dict


def compute_reading_order(shapes_meta: List[Dict]) -> List[Dict]:
    """
    Given shapes_meta with bbox left/top, return shapes sorted top-left -> bottom-right reading order.
    """
    def key_fn(s):
        bbox = s.get("bbox", {})
        return (bbox.get("top", 0), bbox.get("left", 0))
    return sorted(shapes_meta, key=key_fn)


def sanitize_text_for_ppt(text: str) -> str:
    # remove null chars and trim excessive whitespace
    if not text:
        return ""
    return text.replace("\x00", "").strip()
