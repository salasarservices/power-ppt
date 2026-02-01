# preprocessor.py
import io
from typing import Optional

import cv2
import numpy as np
from PIL import Image

# Preprocess image bytes for better OCR accuracy:
# - convert to grayscale
# - denoise (fastNlMeans)
# - adaptive threshold or contrast increase
# - optional upscale

def _read_image_bytes(image_bytes: bytes) -> Optional[np.ndarray]:
    try:
        arr = np.frombuffer(image_bytes, np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        return img
    except Exception:
        return None


def _write_image_bytes(img: np.ndarray) -> bytes:
    success, enc = cv2.imencode(".png", img)
    if not success:
        raise RuntimeError("Failed to encode image")
    return enc.tobytes()


def preprocess_image(image_bytes: bytes, deskew: bool = True, upscale: int = 1, dpi: int = 300) -> bytes:
    """
    Preprocess given image bytes and return processed PNG bytes.
    `upscale` multiplies the image size (1 = no upscale, 2 = 2x).
    """
    img = _read_image_bytes(image_bytes)
    if img is None:
        # try via PIL
        img_pil = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)

    # Convert to grayscale then to color as needed
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Denoise
    den = cv2.fastNlMeansDenoising(gray, None, h=10, templateWindowSize=7, searchWindowSize=21)

    # Improve contrast using CLAHE
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(den)

    # Optionally deskew (simple method using moments)
    if deskew:
        coords = np.column_stack(np.where(enhanced > 0))
        if coords.size != 0:
            angle = cv2.minAreaRect(coords)[-1]
            if angle < -45:
                angle = -(90 + angle)
            else:
                angle = -angle
            (h, w) = enhanced.shape[:2]
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            enhanced = cv2.warpAffine(enhanced, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

    # Upscale
    if upscale and upscale > 1:
        enhanced = cv2.resize(enhanced, None, fx=upscale, fy=upscale, interpolation=cv2.INTER_CUBIC)

    # Merge into BGR for output PNG
    out_bgr = cv2.cvtColor(enhanced, cv2.COLOR_GRAY2BGR)
    return _write_image_bytes(out_bgr)
