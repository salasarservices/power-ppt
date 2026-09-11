from fastapi import APIRouter, File, HTTPException, Query, UploadFile

from ..core.config import get_settings
from ..schemas import AnalyzeResponse
from ..services.analyzer import analyze_pptx

router = APIRouter()


@router.post("/analyze", response_model=AnalyzeResponse)
async def analyze(
    file: UploadFile = File(...),
    use_ocr: bool | None = Query(None, description="Override the server default OCR setting"),
    always_ocr: bool = Query(False, description="OCR every slide, not just text-empty ones"),
):
    """Extract a reviewable SlidePlan from an uploaded PPTX."""
    if not (file.filename or "").lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Please upload a .pptx file.")

    data = await file.read()
    s = get_settings()
    enabled = s.ocr_enabled_default if use_ocr is None else use_ocr

    plan, warnings = analyze_pptx(
        data, use_ocr=enabled, backend=s.ocr_backend, always_ocr=always_ocr, settings=s
    )
    return AnalyzeResponse(slides=len(plan.pages), warnings=warnings, plan=plan)
