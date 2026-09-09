from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .api import analyze, generate, health
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

    app.include_router(health.router, tags=["health"])
    app.include_router(analyze.router, tags=["analyze"])
    app.include_router(generate.router, tags=["generate"])
    return app


app = create_app()
