FROM python:3.12-slim

WORKDIR /app

# System libs for opencv-python-headless (OCR preprocessing).
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 libgl1 tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

# Dependencies (pinned; core + api + ocr). Kept explicit to match pyproject.
RUN pip install --no-cache-dir \
    "python-pptx==1.0.2" "Pillow>=11.1,<13" "pydantic>=2.7,<3" \
    "fastapi>=0.115,<0.117" "uvicorn[standard]>=0.30,<0.35" \
    "pydantic-settings>=2.3,<3" "python-multipart>=0.0.9,<0.1" \
    "opencv-python-headless>=4.9,<5" "numpy>=1.26,<3" \
    "google-cloud-vision>=3.7,<4" "pytesseract>=0.3.10,<0.4" "boto3>=1.34,<2"

COPY app ./app
COPY templates ./templates

ENV PORT=8080
EXPOSE 8080
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
