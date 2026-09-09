import io

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..brand_engine import BrandEngineError, build_deck
from ..core.config import get_settings
from ..schemas import SlidePlan

router = APIRouter()

_PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


@router.post("/generate")
def generate(plan: SlidePlan):
    """Build the standardised deck from a reviewed SlidePlan and stream it back."""
    s = get_settings()
    try:
        data = build_deck(plan, s.template_path)
    except BrandEngineError as e:
        raise HTTPException(status_code=422, detail=str(e))

    return StreamingResponse(
        io.BytesIO(data),
        media_type=_PPTX_MIME,
        headers={"Content-Disposition": 'attachment; filename="standardised.pptx"'},
    )
