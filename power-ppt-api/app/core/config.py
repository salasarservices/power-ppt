"""
Settings — env-driven (prefix POWERPPT_). In production, secret values come from
Google Secret Manager injected as env vars; locally from a .env file.
"""

from functools import lru_cache
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

# power-ppt-api/  (config.py is app/core/config.py -> parents[2])
_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_TEMPLATE = _ROOT / "templates" / "2026" / "Salasar_Corporate_Blank.pptx"


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_prefix="POWERPPT_", env_file=".env", extra="ignore"
    )

    template_path: str = str(_DEFAULT_TEMPLATE)
    template_version: str = "2026"

    # Built SPA served by this service in production (single-service topology).
    # Defaults to /app/web_static inside the image; absent in local/test runs, so
    # the static mount is skipped and the API serves at root as before.
    web_dist_path: str = str(_ROOT / "web_static")

    # Vite dev origin by default; override in prod with the deployed web origin.
    # Unused in single-service prod (same-origin), harmless to keep.
    cors_origins: list[str] = ["http://localhost:5173"]

    # OCR
    ocr_enabled_default: bool = True
    ocr_backend: str = "auto"

    # Google Document AI — image-of-a-table -> native table (A3). Uses the runtime
    # service account (ADC); no key file. Configured = project + processor_id set.
    docai_project: str | None = None
    docai_location: str = "us"          # processor region, e.g. us | eu
    docai_processor_id: str | None = None

    # Secrets (prefer Secret Manager in prod)
    google_service_account_json: str | None = None
    aws_access_key_id: str | None = None
    aws_secret_access_key: str | None = None
    aws_default_region: str = "us-east-1"


@lru_cache
def get_settings() -> Settings:
    return Settings()
