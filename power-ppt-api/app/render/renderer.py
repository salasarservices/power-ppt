"""
Slide rendering via LibreOffice headless. Converts a PPTX to PDF so a vision model
(Vertex AI Gemini) can read slides whose content is baked into pictures/freeform
art. LibreOffice (`soffice`) must be on the image (see Dockerfile / DEPLOY.md).
"""

import os
import shutil
import subprocess
import tempfile


class RenderError(Exception):
    """Raised when a PPTX cannot be rendered (LibreOffice missing or conversion failed)."""


def _find_soffice(explicit: str | None = None) -> str:
    for cand in (explicit, "soffice", "libreoffice"):
        if cand and shutil.which(cand):
            return cand
    raise RenderError(
        "LibreOffice (soffice) not found — install it in the image to enable AI extraction."
    )


def pptx_to_pdf(pptx_bytes: bytes, soffice_bin: str | None = None, timeout: int = 120) -> bytes:
    """Convert PPTX bytes to PDF bytes using LibreOffice headless."""
    soffice = _find_soffice(soffice_bin)
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "in.pptx")
        with open(src, "wb") as f:
            f.write(pptx_bytes)
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, src],
                check=True, capture_output=True, timeout=timeout,
                env={**os.environ, "HOME": tmp},   # LibreOffice needs a writable HOME
            )
        except subprocess.CalledProcessError as e:
            raise RenderError(f"LibreOffice conversion failed: {e.stderr.decode(errors='ignore')[:300]}") from e
        except subprocess.TimeoutExpired as e:
            raise RenderError("LibreOffice conversion timed out.") from e

        out = os.path.join(tmp, "in.pdf")
        if not os.path.exists(out):
            raise RenderError("LibreOffice produced no PDF.")
        with open(out, "rb") as f:
            return f.read()
