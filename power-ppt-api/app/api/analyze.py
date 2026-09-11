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
    use_ai: bool | None = Query(None, description="Override Tier-2 AI (Vertex Gemini) extraction"),
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

    # Tier 2: render + Vertex AI Gemini for slides native extraction can't read.
    ai_on = s.ai_extract_default if use_ai is None else use_ai
    if ai_on:
        from ..ai import gemini
        from ..services.ai_extract import enrich_plan

        if gemini.is_configured(s):
            plan, ai_warnings = enrich_plan(plan, data, s)
            warnings = warnings + ai_warnings
        else:
            warnings.append("AI extraction requested but Vertex AI is not configured.")

    return AnalyzeResponse(slides=len(plan.pages), warnings=warnings, plan=plan)
