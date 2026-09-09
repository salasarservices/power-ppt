import io

from fastapi.testclient import TestClient
from pptx import Presentation

from app.brand_engine import build_deck
from app.main import app
from app.schemas import Page, SlidePlan

client = TestClient(app)


def _sample_pptx() -> bytes:
    plan = SlidePlan(pages=[Page(title="Intro", body="Some body text for the slide.")])
    return build_deck(plan, _template())


def _template() -> str:
    from app.core.config import get_settings

    return get_settings().template_path


def test_health():
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json()["status"] == "ok"


def test_generate_returns_valid_pptx():
    payload = {
        "pages": [
            {"title": "Motor - Overview", "body": "Line one.\n\nLine two.", "tables": []}
        ]
    }
    r = client.post("/generate", json=payload)
    assert r.status_code == 200
    assert "presentationml" in r.headers["content-type"]
    assert r.content[:2] == b"PK"  # zip/OOXML magic
    prs = Presentation(io.BytesIO(r.content))
    assert len(prs.slides) == 1


def test_generate_rejects_empty_plan():
    r = client.post("/generate", json={"pages": []})
    assert r.status_code == 422  # BrandEngineError -> 422


def test_analyze_extracts_a_plan():
    files = {
        "file": (
            "in.pptx",
            io.BytesIO(_sample_pptx()),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    }
    r = client.post("/analyze?use_ocr=false", files=files)
    assert r.status_code == 200
    data = r.json()
    assert data["slides"] >= 1
    assert "pages" in data["plan"]
    joined = " ".join(p["body"] for p in data["plan"]["pages"])
    assert "body text" in joined.lower()


def test_analyze_rejects_non_pptx():
    files = {"file": ("notes.txt", io.BytesIO(b"hello"), "text/plain")}
    r = client.post("/analyze", files=files)
    assert r.status_code == 400
