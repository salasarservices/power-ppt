from fastapi import APIRouter

from ..core.config import get_settings

router = APIRouter()


@router.get("/health")
def health():
    s = get_settings()
    return {"status": "ok", "template_version": s.template_version}
