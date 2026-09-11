"""
Data contract shared by every content source (reformat mode today, v2 AI Generate
mode later) and consumed by the brand engine. The engine only ever sees a SlidePlan.
"""

from typing import Literal

from pydantic import BaseModel, Field


class Table(BaseModel):
    header: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)  # includes the header as rows[0]


class Image(BaseModel):
    data: str                       # base64-encoded image bytes
    content_type: str = "image/png"


class Page(BaseModel):
    title: str = ""          # "" -> the engine renders no heading
    body: str = ""           # paragraphs separated by "\n\n"
    tables: list[Table] = Field(default_factory=list)
    images: list[Image] = Field(default_factory=list)   # source images, auto-fit in the content zone


class SlidePlan(BaseModel):
    pages: list[Page] = Field(default_factory=list)


class AnalyzeResponse(BaseModel):
    slides: int
    warnings: list[str] = Field(default_factory=list)
    plan: SlidePlan


# ── Layout (Placement JSON) — the positioned output of the flow engine, and the
# editor's round-trip contract: /layout produces it, the canvas renders + edits it,
# /render turns it back into the exact .pptx. Positions are in inches. ──────────
class Placement(BaseModel):
    kind: Literal["body", "table", "image"]
    left: float
    top: float
    width: float
    height: float | None = None      # body only; table/image derive their own height
    text: str | None = None          # body
    table: Table | None = None       # table
    image: Image | None = None       # image


class Slide(BaseModel):
    title: str = ""
    placements: list[Placement] = Field(default_factory=list)


class Deck(BaseModel):
    slides: list[Slide] = Field(default_factory=list)
