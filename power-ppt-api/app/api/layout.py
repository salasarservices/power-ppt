from fastapi import APIRouter

from ..brand_engine import flow_deck
from ..schemas import Deck, SlidePlan

router = APIRouter()


@router.post("/layout", response_model=Deck)
def layout(plan: SlidePlan):
    """Flow a reviewed SlidePlan into a positioned Deck (Placement JSON) for the
    editor canvas."""
    return flow_deck(plan.pages)
