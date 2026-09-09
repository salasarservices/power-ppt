from pathlib import Path

import pytest

from app.brand_engine import BrandEngineError, build_deck, verify_output
from app.schemas import Page, SlidePlan

TEMPLATE = str(
    Path(__file__).parents[1] / "templates" / "2026" / "Salasar_Corporate_Blank.pptx"
)


def test_verify_rejects_non_pptx():
    with pytest.raises(BrandEngineError):
        verify_output(b"this is not a pptx", [{"title": "x", "body": "", "tables": []}], TEMPLATE)


def test_verify_rejects_wrong_slide_count():
    data = build_deck(SlidePlan(pages=[Page(title="One", body="Body.")]), TEMPLATE)
    # Claim two pages were rendered when only one was.
    with pytest.raises(BrandEngineError):
        verify_output(
            data,
            [
                {"title": "One", "body": "Body.", "tables": []},
                {"title": "Two", "body": "Body.", "tables": []},
            ],
            TEMPLATE,
        )


def test_build_deck_passes_its_own_guards():
    # Should not raise — build_deck runs verify_output internally.
    data = build_deck(SlidePlan(pages=[Page(title="Fine", body="Body.")]), TEMPLATE)
    assert isinstance(data, bytes) and len(data) > 0
