"""
Turn an uploaded PPTX into a SlidePlan: extract editable title/body/tables, carry
source images, and (optionally) OCR. Two OCR roles:
  - Document AI (A3): an image that is actually a table becomes a native table and
    the image is dropped.
  - Tesseract/cloud text OCR: pull text off an otherwise text-empty image slide.
OCR dependencies are imported lazily so the API runs even when they are not
installed — a missing OCR stack degrades to a warning, never a crash.
"""

import base64

from ..core.config import Settings
from ..schemas import Image, Page, SlidePlan, Table


def _to_image(img) -> Image | None:
    data = img.get("image_bytes")
    if not data:
        return None
    return Image(
        data=base64.b64encode(data).decode("ascii"),
        content_type=img.get("content_type") or "image/png",
    )


def _native_tables(meta) -> list[Table]:
    tables: list[Table] = []
    for shape in meta.get("shapes", []):
        try:
            if getattr(shape, "has_table", False) and shape.has_table:
                rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
                if rows:
                    tables.append(Table(header=rows[0], rows=rows))
        except Exception:
            pass
    return tables


def analyze_pptx(
    pptx_bytes: bytes,
    use_ocr: bool = False,
    backend: str = "auto",
    always_ocr: bool = False,
    settings: Settings | None = None,
) -> tuple[SlidePlan, list[str]]:
    from ..extract import pptx_reader

    slides_meta = pptx_reader.extract_text_shapes(pptx_bytes)
    warnings: list[str] = []

    ocr_image = preprocess = None
    if use_ocr:
        try:
            from ..ocr import ocr_backend as _ob
            from ..ocr import preprocessor as _pp

            ocr_image, preprocess = _ob.ocr_image, _pp.preprocess_image
        except Exception as e:  # OCR deps not installed / import failure
            warnings.append(f"OCR unavailable ({e}); proceeding without OCR.")
            use_ocr = False

    # Document AI (A3) — image-of-a-table -> native table.
    docai = None
    if use_ocr and settings is not None:
        from ..ocr import docai as _docai

        if _docai.is_configured(settings):
            docai = _docai

    pages: list[Page] = []
    for meta in slides_meta:
        idx = meta.get("slide_index", 0)
        title = meta.get("title_text") or ""
        body = meta.get("body_text", "") or ""
        tables = _native_tables(meta)

        # First pass: convert any image that is really a table; keep the rest.
        kept_images: list[dict] = []
        for img in meta.get("image_shapes", []):
            consumed = False
            if docai is not None:
                try:
                    found = docai.image_to_tables(
                        img["image_bytes"], img.get("content_type"), settings
                    )
                    if found:
                        tables.extend(found)
                        consumed = True
                except Exception as e:
                    warnings.append(f"Slide {idx + 1}: table OCR failed ({e}).")
            if not consumed:
                kept_images.append(img)

        # Second pass: text OCR on a text-empty slide's remaining images.
        if use_ocr and (not body.strip() or always_ocr):
            for img in kept_images:
                try:
                    processed = preprocess(img["image_bytes"])
                    result = ocr_image(processed, backend=backend)
                    text = (result.get("text") or "").strip()
                    if text:
                        body = (body + "\n\n" + text).strip() if body else text
                    for t in result.get("tables", []):
                        if t.get("rows"):
                            tables.append(Table(header=t.get("header", []), rows=t["rows"]))
                except Exception as e:
                    warnings.append(f"Slide {idx + 1}: OCR failed ({e}).")

        images = [im for im in (_to_image(i) for i in kept_images) if im is not None]
        pages.append(Page(title=title, body=body, tables=tables, images=images))

    return SlidePlan(pages=pages), warnings
