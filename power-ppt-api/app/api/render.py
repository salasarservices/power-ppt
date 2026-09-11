import io

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from ..brand_engine import BrandEngineError, render_deck
from ..core.config import get_settings
from ..schemas import Deck

router = APIRouter()

_PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


@router.post("/render")
def render(deck: Deck):
    """Render the reviewed/edited Deck (Placement JSON) into the standardised
    .pptx — the exact placements the canvas showed (WYSIWYG)."""
    s = get_settings()
    try:
        data = render_deck(deck, s.template_path)
    except BrandEngineError as e:
        raise HTTPException(status_code=422, detail=str(e))

    return StreamingResponse(
        io.BytesIO(data),
        media_type=_PPTX_MIME,
        headers={"Content-Disposition": 'attachment; filename="standardised.pptx"'},
    )
