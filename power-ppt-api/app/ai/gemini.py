"""
Vision extraction via Vertex AI Gemini (Tier 2). Reads a rendered deck (PDF) and
returns structured, editable content per slide — recovering text and tables that
are baked into pictures / freeform art, which native extraction cannot see.

COMPLIANCE (non-negotiable): Vertex AI only (enterprise, no-train). Never the
free-tier / AI Studio endpoint. The model must transcribe what is visibly present
and never invent coverage, figures, dates or statistics (IRDAI + client-data rules).
"""

import json

from ..core.config import Settings
from ..schemas import Page, SlidePlan, Table

_PROMPT = (
    "You are transcribing a slide deck into structured JSON for reformatting. "
    "For EACH slide in the attached PDF, in order, extract only what is VISIBLY "
    "present — transcribe text and tables verbatim. Do NOT summarise, translate, "
    "infer, or invent any text, numbers, dates, coverage, or statistics; if "
    "something is unreadable, omit it. Return STRICT JSON of the form: "
    '{"slides":[{"title":"<slide title or empty>","body":"<paragraphs separated by '
    'blank lines>","tables":[{"header":["c1","c2"],"rows":[["c1","c2"],["v1","v2"]]}]}]}. '
    "One array item per slide. Body excludes the title and any repeated footer/tagline."
)


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


def extract_raw(pdf_bytes: bytes, s: Settings) -> str:
    """Send the rendered deck PDF to Vertex AI Gemini; return the raw JSON text."""
    from google import genai            # google-genai SDK, lazy
    from google.genai import types

    client = genai.Client(vertexai=True, project=s.vertex_project, location=s.vertex_location)
    resp = client.models.generate_content(
        model=s.vertex_model,
        contents=[
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            _PROMPT,
        ],
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0,
            max_output_tokens=32768,
        ),
    )
    return resp.text or "{}"


def extract_pages(pdf_bytes: bytes, s: Settings) -> list[Page]:
    """Structured pages from Vertex AI Gemini. Raises on library/API errors so the
    caller can warn and fall back to native."""
    return parse_response(extract_raw(pdf_bytes, s))


def extract_plan(pdf_bytes: bytes, s: Settings) -> SlidePlan:
    return SlidePlan(pages=extract_pages(pdf_bytes, s))
