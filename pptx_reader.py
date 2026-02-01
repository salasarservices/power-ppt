# pptx_reader.py
import io
from typing import List, Dict, Any

from pptx import Presentation
from pptx.util import Emu

# Utility to extract shapes and embedded pictures from slides

def _get_shape_bbox(shape) -> Dict[str, int]:
    """
    Return bounding box of a shape in EMU units
    """
    try:
        return {"left": int(shape.left), "top": int(shape.top), "width": int(shape.width), "height": int(shape.height)}
    except Exception:
        return {"left": 0, "top": 0, "width": 0, "height": 0}


def extract_text_shapes(presentation_bytes: bytes) -> List[Dict[str, Any]]:
    """
    Parse presentation bytes and return meta per slide including:
    - slide_index
    - title_text (if found via placeholder)
    - text_shapes: list of {'bbox':..., 'text': str}
    - image_shapes: list of {'bbox':..., 'image_bytes': bytes}
    """
    prs = Presentation(io.BytesIO(presentation_bytes))
    slides_meta = []
    for idx, slide in enumerate(prs.slides):
        meta = {"slide_index": idx, "text_shapes": [], "image_shapes": [], "title_text": None}

        # Try to find title placeholder
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            try:
                ph = shape.placeholder_format
                if ph.type.name in ("TITLE", "CENTER_TITLE"):
                    txt = shape.text.strip()
                    if txt:
                        meta["title_text"] = txt
            except Exception:
                # Not all shapes expose placeholder_format reliably
                pass

        # Collect text shapes and images
        for shape in slide.shapes:
            # text shapes
            if getattr(shape, "has_text_frame", False):
                txt = shape.text.strip()
                if txt:
                    bbox = _get_shape_bbox(shape)
                    meta["text_shapes"].append({"bbox": bbox, "text": txt})
            # image shapes (pictures)
            try:
                if hasattr(shape, "image") and shape.image is not None:
                    image_bytes = shape.image.blob
                    bbox = _get_shape_bbox(shape)
                    meta["image_shapes"].append({"bbox": bbox, "image_bytes": image_bytes})
            except Exception:
                # some shapes may throw; skip them
                pass

        slides_meta.append(meta)
    return slides_meta
