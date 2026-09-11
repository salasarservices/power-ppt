# Single-service image for Cloud Run: builds the Vite SPA, then bundles it into
# the FastAPI runtime so one container serves both the UI and the API.
# Build context = repo root.  Build:  gcloud run deploy --source .

# ---- stage 1: build the SPA ----
FROM node:20-slim AS web
WORKDIR /web
COPY power-ppt-web/package.json power-ppt-web/package-lock.json ./
RUN npm ci
COPY power-ppt-web/ ./
RUN npm run build   # -> /web/dist

# ---- stage 2: API runtime + bundled SPA ----
FROM python:3.12-slim
WORKDIR /app

# System libs: opencv (OCR preprocessing) + tesseract; LibreOffice Impress renders
# PPTX->PDF for Tier-2 Vertex AI Gemini extraction of picture/freeform decks.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 libgl1 tesseract-ocr libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*

# Dependencies (pinned; core + api + ocr). Kept explicit to match pyproject.
RUN pip install --no-cache-dir \
    "python-pptx==1.0.2" "Pillow>=11.1,<13" "pydantic>=2.7,<3" \
    "fastapi>=0.115,<0.117" "uvicorn[standard]>=0.30,<0.35" \
    "pydantic-settings>=2.3,<3" "python-multipart>=0.0.9,<0.1" \
    "opencv-python-headless>=4.9,<5" "numpy>=1.26,<3" \
    "google-cloud-vision>=3.7,<4" "google-cloud-documentai>=2.20,<4" \
    "pytesseract>=0.3.10,<0.4" "boto3>=1.34,<2" "google-genai>=1.0,<2"

COPY power-ppt-api/app ./app
COPY power-ppt-api/templates ./templates
# Bundle the built SPA; main.py mounts it at / when this dir exists.
COPY --from=web /web/dist ./web_static

# Cloud Run injects PORT (defaults to 8080); web_static path matches config default.
ENV PORT=8080
EXPOSE 8080
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
