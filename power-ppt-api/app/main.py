from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .api import aidebug, analyze, generate, health, layout, render
from .core.config import get_settings


def create_app() -> FastAPI:
    s = get_settings()
    app = FastAPI(title="PowerPPT API", version="0.1.0")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=s.cors_origins,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # API routes are registered first so they always win over the SPA catch-all.
    app.include_router(health.router, tags=["health"])
    app.include_router(analyze.router, tags=["analyze"])
    app.include_router(generate.router, tags=["generate"])
    app.include_router(layout.router, tags=["layout"])
    app.include_router(render.router, tags=["render"])
    app.include_router(aidebug.router, tags=["debug"])

    # Single-service topology: if a built SPA is bundled (in the Cloud Run image),
    # serve it at root. Skipped in local/test runs where the dir doesn't exist.
    web_dist = Path(s.web_dist_path)
    if web_dist.is_dir():
        app.mount("/", StaticFiles(directory=str(web_dist), html=True), name="web")

    return app


app = create_app()
