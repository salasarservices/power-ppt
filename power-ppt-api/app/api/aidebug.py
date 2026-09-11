"""Temporary diagnostic endpoint: expose the render + raw Gemini output so we can
see exactly where content is lost. Remove once the extraction is dialled in."""

import io

from fastapi import APIRouter, File, Query, UploadFile
from fastapi.responses import StreamingResponse

from ..ai import gemini
from ..core.config import get_settings
from ..render import pptx_to_pdf

router = APIRouter()


@router.post("/ai-debug")
async def ai_debug(file: UploadFile = File(...), pdf: bool = Query(False)):
    data = await file.read()
    s = get_settings()

    if pdf:  # return the LibreOffice render so we can eyeball it
        rendered = pptx_to_pdf(data, s.soffice_bin)
        return StreamingResponse(io.BytesIO(rendered), media_type="application/pdf")

    out: dict = {"configured": gemini.is_configured(s), "model": s.vertex_model}

    try:
        pdf = pptx_to_pdf(data, s.soffice_bin)
        out["pdf_bytes"] = len(pdf)
    except Exception as e:
        out["render_error"] = str(e)
        return out

    try:
        raw = gemini.extract_raw(pdf, s)
        out["raw_len"] = len(raw)
        out["raw"] = raw[:40000]
        out["parsed_slides"] = len(gemini.parse_response(raw))
    except Exception as e:
        out["ai_error"] = str(e)
    return out
