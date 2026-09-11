"""
Turn an uploaded PPTX into a SlidePlan: extract editable title/body/tables, and
(optionally) OCR image-only slides. OCR dependencies are imported lazily so the API
runs even when they are not installed — a missing OCR stack degrades to a warning,
never a crash (roadmap: surface OCR failures, don't swallow them).
"""

import base64

from ..schemas import Image, Page, SlidePlan, Table


def _source_images(meta) -> list[Image]:
    """Carry the slide's source images into the plan (base64) so the brand engine
    re-places them in the content zone."""
    out: list[Image] = []
    for img in meta.get("image_shapes", []):
        data = img.get("image_bytes")
        if not data:
            continue
        out.append(
            Image(
                data=base64.b64encode(data).decode("ascii"),
                content_type=img.get("content_type") or "image/png",
            )
        )
    return out


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

    pages: list[Page] = []
    for meta in slides_meta:
        idx = meta.get("slide_index", 0)
        title = meta.get("title_text") or ""
        body = meta.get("body_text", "") or ""
        tables = _native_tables(meta)

        if use_ocr and (not body.strip() or always_ocr):
            for img in meta.get("image_shapes", []):
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

        pages.append(Page(title=title, body=body, tables=tables, images=_source_images(meta)))

    return SlidePlan(pages=pages), warnings
