"""
Tier-2 reconciler. Native extraction (Tier 1) is lossless and editable, so it wins
wherever it recovers real content. For slides where native comes up short — the
"designed" slides whose text/tables are baked into pictures/freeform art — we render
the deck and let Vertex AI Gemini transcribe them, then merge, always keeping the
native source images so figures/charts are preserved.
"""

from ..ai import gemini
from ..core.config import Settings
from ..render import pptx_to_pdf
from ..schemas import Page, SlidePlan


def _native_sufficient(p: Page) -> bool:
    """True when native extraction already recovered real content for this slide."""
    return bool(p.title) or len(p.body.strip()) >= 40 or bool(p.tables)


def enrich_plan(native: SlidePlan, pptx_bytes: bytes, s: Settings) -> tuple[SlidePlan, list[str]]:
    """Fill in content-poor slides with Vertex AI Gemini transcription. Falls back
    to native (never crashes) on any render/AI failure or slide-count mismatch."""
    warnings: list[str] = []
    try:
        pdf = pptx_to_pdf(pptx_bytes, s.soffice_bin)
        ai_pages = gemini.extract_pages(pdf, s)
    except Exception as e:
        warnings.append(f"AI extraction unavailable ({e}); used native extraction only.")
        return native, warnings

    if len(ai_pages) != len(native.pages):
        warnings.append(
            f"AI returned {len(ai_pages)} slides vs {len(native.pages)} native; used native only."
        )
        return native, warnings

    merged: list[Page] = []
    for np, ap in zip(native.pages, ai_pages):
        if _native_sufficient(np):
            merged.append(np)                       # native already carries its images
        else:
            merged.append(Page(
                title=ap.title or np.title,
                body=ap.body,
                tables=ap.tables or np.tables,
                images=np.images,                   # keep source figures/charts
            ))
    return SlidePlan(pages=merged), warnings
