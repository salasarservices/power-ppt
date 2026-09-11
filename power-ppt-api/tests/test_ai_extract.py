from app.ai import gemini
from app.core.config import Settings
from app.schemas import Image, Page, SlidePlan
from app.services import ai_extract


# ── Gemini response parsing (pure) ───────────────────────────────────────────
def test_parse_response_slides_with_table():
    text = (
        '{"slides":[{"title":"Premium","body":"Line one.","tables":'
        '[{"header":["Item","Amt"],"rows":[["Item","Amt"],["OD","8500"]]}]}]}'
    )
    pages = gemini.parse_response(text)
    assert len(pages) == 1
    assert pages[0].title == "Premium"
    assert pages[0].tables[0].rows == [["Item", "Amt"], ["OD", "8500"]]


def test_parse_response_accepts_bare_list_and_missing_keys():
    pages = gemini.parse_response('[{"body":"just text"}]')
    assert len(pages) == 1
    assert pages[0].title == "" and pages[0].body == "just text"


def test_is_configured():
    assert not gemini.is_configured(Settings())
    assert gemini.is_configured(Settings(vertex_project="p"))


# ── Reconciler ───────────────────────────────────────────────────────────────
def _native_sufficient_page():
    return Page(title="Real Title", body="", tables=[])


def _native_poor_page():
    return Page(title="", body="", tables=[], images=[Image(data="Zm9v")])


def test_native_sufficient_heuristic():
    assert ai_extract._native_sufficient(_native_sufficient_page())
    assert not ai_extract._native_sufficient(_native_poor_page())


def test_enrich_replaces_only_poor_slides_and_keeps_images(monkeypatch):
    native = SlidePlan(pages=[_native_sufficient_page(), _native_poor_page()])
    ai_pages = [Page(title="AI1", body="ai body 1"), Page(title="AI2", body="ai body 2")]

    monkeypatch.setattr(ai_extract, "pptx_to_pdf", lambda *a, **k: b"%PDF-")
    monkeypatch.setattr(ai_extract.gemini, "extract_pages", lambda *a, **k: ai_pages)

    plan, warnings = ai_extract.enrich_plan(native, b"pptx", Settings(vertex_project="p"))
    assert warnings == []
    assert plan.pages[0].title == "Real Title"          # sufficient -> native kept
    assert plan.pages[1].body == "ai body 2"            # poor -> AI content used
    assert plan.pages[1].images == native.pages[1].images  # native images preserved


def test_enrich_falls_back_on_render_error(monkeypatch):
    native = SlidePlan(pages=[_native_poor_page()])

    def boom(*a, **k):
        raise RuntimeError("no libreoffice")

    monkeypatch.setattr(ai_extract, "pptx_to_pdf", boom)
    plan, warnings = ai_extract.enrich_plan(native, b"pptx", Settings(vertex_project="p"))
    assert plan is native
    assert warnings and "used native" in warnings[0]


def test_enrich_falls_back_on_slide_count_mismatch(monkeypatch):
    native = SlidePlan(pages=[_native_poor_page(), _native_poor_page()])
    monkeypatch.setattr(ai_extract, "pptx_to_pdf", lambda *a, **k: b"%PDF-")
    monkeypatch.setattr(ai_extract.gemini, "extract_pages", lambda *a, **k: [Page(body="one")])

    plan, warnings = ai_extract.enrich_plan(native, b"pptx", Settings(vertex_project="p"))
    assert plan is native
    assert warnings and "vs" in warnings[0]
