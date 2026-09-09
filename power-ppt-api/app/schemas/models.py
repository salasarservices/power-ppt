"""
Data contract shared by every content source (reformat mode today, v2 AI Generate
mode later) and consumed by the brand engine. The engine only ever sees a SlidePlan.
"""

from pydantic import BaseModel, Field


class Table(BaseModel):
    header: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)  # includes the header as rows[0]


class Page(BaseModel):
    title: str = ""          # "" -> the engine renders no heading
    body: str = ""           # paragraphs separated by "\n\n"
    tables: list[Table] = Field(default_factory=list)


class SlidePlan(BaseModel):
    pages: list[Page] = Field(default_factory=list)


class AnalyzeResponse(BaseModel):
    slides: int
    warnings: list[str] = Field(default_factory=list)
    plan: SlidePlan
