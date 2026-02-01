# ocr_backend.py
import io
import os
from typing import Optional, Dict

from PIL import Image

google_client = None

def init_google_vision(credentials_json: Optional[str] = None):
    """
    Initialize Google Cloud Vision client.
    - credentials_json: full service account JSON string (optional). If None, rely on environment or ADC.
    """
    global google_client
    try:
        from google.cloud import vision
        from google.oauth2 import service_account
    except Exception as e:
        raise RuntimeError("google-cloud-vision package is required for Google Vision integration") from e

    if credentials_json:
        creds = service_account.Credentials.from_service_account_info(__import__("json").loads(credentials_json))
        google_client = vision.ImageAnnotatorClient(credentials=creds)
    else:
        # ADC or environment based
        google_client = vision.ImageAnnotatorClient()


def _google_vision_ocr_bytes(image_bytes: bytes) -> Dict:
    """
    Return {'text': full_text, 'lines': [...]} using Google Vision API.
    """  
    if google_client is None:
        raise RuntimeError("Google Vision client not initialized")

    from google.cloud import vision  # local import

    image = vision.Image(content=image_bytes)
    response = google_client.document_text_detection(image=image)
    if response.error.message:
        raise RuntimeError(response.error.message)
    full_text = response.full_text_annotation.text if response.full_text_annotation is not None else ""
    lines = []
    # document_text_detection gives structured pages/blocks/paragraphs/words; we gather lines
    if response.full_text_annotation:
        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    line = []
                    for word in paragraph.words:
                        word_text = "".join([symbol.text for symbol in word.symbols])
                        line.append(word_text)
                    if line:
                        lines.append(" ".join(line))
    return {"text": full_text, "lines": lines}

# Simple Tesseract fallback
def _tesseract_ocr_bytes(image_bytes: bytes) -> Dict:
    try:
        import pytesseract
    except Exception as e:
        raise RuntimeError("pytesseract is required for tesseract OCR") from e

    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    text = pytesseract.image_to_string(img)
    lines = [l for l in text.splitlines() if l.strip()]
    return {"text": text, "lines": lines}


def ocr_image(image_bytes: bytes, backend: str = "google_vision") -> Dict:
    """
    Generic OCR wrapper. Returns dict with 'text' and 'lines'.
    backend: 'google_vision' or 'tesseract'
    """  
    if backend == "google_vision":
        try:
            return _google_vision_ocr_bytes(image_bytes)
        except Exception as e:
            # Fall back to tesseract if available
            try:
                return _tesseract_ocr_bytes(image_bytes)
            except Exception:
                raise e
    elif backend == "tesseract":
        return _tesseract_ocr_bytes(image_bytes)
    else:
        raise ValueError(f"Unknown OCR backend: {backend}")
