"""
Google Document AI — convert an image that is actually a table into a native,
editable table (roadmap A3). Uses a Form Parser processor; auth is the runtime
service account (ADC), so no key file. The client library is imported lazily so
the app runs without it installed.

Setup (once, per DEPLOY.md): enable documentai.googleapis.com, create a Form
Parser processor, grant the Cloud Run service account roles/documentai.apiUser,
and set POWERPPT_DOCAI_PROJECT / _LOCATION / _PROCESSOR_ID.
"""

from ..core.config import Settings
from ..schemas import Table


def is_configured(s: Settings) -> bool:
    return bool(s.docai_project and s.docai_processor_id)


def _cell_text(layout, full_text: str) -> str:
    """Resolve a Document AI cell's text from its text-anchor segments."""
    anchor = getattr(layout, "text_anchor", None)
    if anchor is None:
        return ""
    parts = []
    for seg in anchor.text_segments:
        start = int(getattr(seg, "start_index", 0) or 0)
        end = int(getattr(seg, "end_index", 0) or 0)
        parts.append(full_text[start:end])
    return "".join(parts).strip()


def _row_cells(row, full_text: str) -> list[str]:
    return [_cell_text(cell.layout, full_text) for cell in row.cells]


def parse_tables(document) -> list[Table]:
    """Pure parser (testable without the API): Document -> list[Table]. Each table's
    rows[0] is the header (repeated in `header`)."""
    full_text = document.text or ""
    out: list[Table] = []
    for page in document.pages:
        for table in page.tables:
            rows: list[list[str]] = []
            for hr in table.header_rows:
                rows.append(_row_cells(hr, full_text))
            for br in table.body_rows:
                rows.append(_row_cells(br, full_text))
            if rows:
                out.append(Table(header=rows[0], rows=rows))
    return out


def image_to_tables(image_bytes: bytes, content_type: str, s: Settings) -> list[Table]:
    """Process an image through Document AI; return any tables found (empty list if
    the image is not a table). Raises on API/library errors so the caller can warn."""
    from google.cloud import documentai_v1 as documentai  # lazy

    client = documentai.DocumentProcessorServiceClient(
        client_options={"api_endpoint": f"{s.docai_location}-documentai.googleapis.com"}
    )
    name = client.processor_path(s.docai_project, s.docai_location, s.docai_processor_id)
    raw = documentai.RawDocument(
        content=image_bytes, mime_type=content_type or "image/png"
    )
    result = client.process_document(
        request=documentai.ProcessRequest(name=name, raw_document=raw)
    )
    return parse_tables(result.document)
