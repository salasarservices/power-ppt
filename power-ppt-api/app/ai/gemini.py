"""
Vision extraction via Vertex AI Gemini (Tier 2). Reads a rendered deck (PDF) and
returns structured, editable content per slide — recovering text and tables that
are baked into pictures / freeform art, which native extraction cannot see.

COMPLIANCE (non-negotiable): Vertex AI only (enterprise, no-train). Never the
free-tier / AI Studio endpoint. The model must transcribe what is visibly present
and never invent coverage, figures, dates or statistics (IRDAI + client-data rules).
"""

import json
from concurrent.futures import ThreadPoolExecutor

from ..core.config import Settings
from ..render import split_pdf_pages
from ..schemas import Page, SlidePlan, Table

# Whole-deck prompt (used by /ai-debug only).
_PROMPT = (
    "You are transcribing a slide deck into structured JSON for reformatting. "
    "For EACH slide in the attached PDF, in order, extract only what is VISIBLY "
    "present — transcribe text and tables verbatim. Do NOT summarise, translate, "
    "infer, or invent. Return STRICT JSON: "
    '{"slides":[{"title":"","body":"","tables":[{"header":[],"rows":[]}]}]}.'
)

# Per-slide prompt — one page at a time gives the model full attention, so it reads
# the real content instead of anchoring on the repeated footer tagline.
_SLIDE_PROMPT = (
    "This image is ONE slide. Transcribe ALL visible text VERBATIM — the title, "
    "every heading, bullet, label, caption, figure and number, and every table with "
    "its rows and columns. Do NOT summarise, translate, infer, or invent anything; "
    "transcribe only what is visibly present. IGNORE the repeated footer tagline "
    "'YOU MANAGE YOUR ENTERPRISE / WE WILL MANAGE YOUR INSURABLE RISKS' and slide "
    'numbers. Return STRICT JSON: {"title":"<slide title or empty>","body":'
    '"<all other text, paragraphs separated by blank lines>","tables":'
    '[{"header":["c1"],"rows":[["c1"],["v1"]]}]}.'
)

_MAX_WORKERS = 8


def is_configured(s: Settings) -> bool:
    return bool(s.vertex_project and s.vertex_model)


def parse_response(text: str) -> list[Page]:
    """Pure parser (testable without Vertex): model JSON -> list[Page]."""
    data = json.loads(text)
    slides = data.get("slides", data) if isinstance(data, dict) else data
    pages: list[Page] = []
    for s in slides or []:
        if not isinstance(s, dict):
            continue
        tables: list[Table] = []
        for t in s.get("tables") or []:
            rows = t.get("rows") or []
            if rows:
                header = t.get("header") or rows[0]
                tables.append(Table(header=header, rows=rows))
        pages.append(Page(
            title=(s.get("title") or "").strip(),
            body=(s.get("body") or "").strip(),
            tables=tables,
        ))
    return pages


def _client(s: Settings):
    from google import genai            # google-genai SDK, lazy
    return genai.Client(vertexai=True, project=s.vertex_project, location=s.vertex_location)


def _generate(client, model: str, pdf_bytes: bytes, prompt: str, max_tokens: int) -> str:
    from google.genai import types
    resp = client.models.generate_content(
        model=model,
        contents=[types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"), prompt],
        config=types.GenerateContentConfig(
            response_mime_type="application/json", temperature=0, max_output_tokens=max_tokens,
        ),
    )
    return resp.text or "{}"


def _parse_page(text: str) -> Page:
    """Parse one slide's JSON object into a Page (tolerant of shape)."""
    try:
        d = json.loads(text)
    except Exception:
        return Page()
    if isinstance(d, list):
        d = d[0] if d else {}
    if not isinstance(d, dict):
        return Page()
    tables: list[Table] = []
    for t in d.get("tables") or []:
        rows = t.get("rows") or []
        if rows:
            tables.append(Table(header=t.get("header") or rows[0], rows=rows))
    return Page(title=(d.get("title") or "").strip(), body=(d.get("body") or "").strip(), tables=tables)


def extract_raw(pdf_bytes: bytes, s: Settings) -> str:
    """Whole-deck transcription (used by /ai-debug only)."""
    return _generate(_client(s), s.vertex_model, pdf_bytes, _PROMPT, 32768)


def extract_pages(pdf_bytes: bytes, s: Settings) -> list[Page]:
    """Per-slide transcription: split the PDF and read each page on its own (in
    parallel) so the model reads real content, not just the repeated tagline.
    Raises on library/API errors so the caller can warn and fall back to native."""
    page_pdfs = split_pdf_pages(pdf_bytes)
    client = _client(s)

    def work(one: bytes) -> Page:
        return _parse_page(_generate(client, s.vertex_model, one, _SLIDE_PROMPT, 8192))

    with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as ex:
        return list(ex.map(work, page_pdfs))


def extract_plan(pdf_bytes: bytes, s: Settings) -> SlidePlan:
    return SlidePlan(pages=extract_pages(pdf_bytes, s))
