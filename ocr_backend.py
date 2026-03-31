# ocr_backend.py
import io
import os
from typing import Optional, Dict, List

import cv2
import numpy as np
from PIL import Image

google_client = None
textract_client = None


# ---------------------------------------------------------------------------
# Client initializers
# ---------------------------------------------------------------------------

def init_google_vision(credentials_json: Optional[str] = None):
    """
    Initialize Google Cloud Vision client.
    credentials_json: full service account JSON string (optional).
    If None, relies on environment / ADC.
    """
    global google_client
    try:
        from google.cloud import vision
        from google.oauth2 import service_account
    except Exception as e:
        raise RuntimeError("google-cloud-vision package is required for Google Vision integration") from e

    if credentials_json:
        creds = service_account.Credentials.from_service_account_info(
            __import__("json").loads(credentials_json)
        )
        google_client = vision.ImageAnnotatorClient(credentials=creds)
    else:
        google_client = vision.ImageAnnotatorClient()


def init_textract(
    aws_access_key_id: Optional[str] = None,
    aws_secret_access_key: Optional[str] = None,
    region_name: str = "us-east-1",
):
    """
    Initialize AWS Textract client.
    If key/secret are omitted, boto3 falls back to environment variables
    (AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY) or an IAM role.
    """
    global textract_client
    try:
        import boto3
    except ImportError as e:
        raise RuntimeError("boto3 is required for AWS Textract integration. Run: pip install boto3") from e

    kwargs: Dict = {"region_name": region_name}
    if aws_access_key_id and aws_secret_access_key:
        kwargs["aws_access_key_id"] = aws_access_key_id
        kwargs["aws_secret_access_key"] = aws_secret_access_key

    textract_client = boto3.client("textract", **kwargs)


# ---------------------------------------------------------------------------
# Table detection heuristic (OpenCV)
# ---------------------------------------------------------------------------

def _has_table_heuristic(image_bytes: bytes, min_lines: int = 3) -> bool:
    """
    Use morphological line detection to decide if an image likely contains a table.
    Looks for at least `min_lines` horizontal AND vertical lines spanning >= 10% of
    the image dimension — the signature of a grid/table structure.
    Returns True if a table is detected.
    """
    arr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_GRAYSCALE)
    if img is None:
        return False

    h, w = img.shape
    _, thresh = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Detect horizontal lines (long, thin kernel)
    h_kernel_len = max(w // 10, 30)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_kernel_len, 1))
    h_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel)

    # Detect vertical lines (tall, thin kernel)
    v_kernel_len = max(h // 10, 30)
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_kernel_len))
    v_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel)

    # Require enough pixels to indicate at least min_lines of each orientation
    h_present = cv2.countNonZero(h_lines) > w * min_lines
    v_present = cv2.countNonZero(v_lines) > h * min_lines

    return h_present and v_present


# ---------------------------------------------------------------------------
# Textract response parser
# ---------------------------------------------------------------------------

def _parse_textract_blocks(blocks: List[Dict]) -> Dict:
    """
    Convert raw Textract Blocks into {'text', 'lines', 'tables'}.

    LINE blocks  → lines list (also joined into full text).
    TABLE blocks → list of {'header': [...], 'rows': [[...], ...]}.
    """
    block_map = {b["Id"]: b for b in blocks}
    lines: List[str] = []
    tables: List[Dict] = []

    for block in blocks:
        btype = block["BlockType"]

        if btype == "LINE":
            lines.append(block.get("Text", ""))

        elif btype == "TABLE":
            cells: Dict = {}
            max_row, max_col = 0, 0

            for rel in block.get("Relationships", []):
                if rel["Type"] != "CHILD":
                    continue
                for cell_id in rel["Ids"]:
                    cell = block_map.get(cell_id)
                    if not cell or cell["BlockType"] not in ("CELL", "MERGED_CELL"):
                        continue
                    row, col = cell["RowIndex"], cell["ColumnIndex"]
                    max_row = max(max_row, row)
                    max_col = max(max_col, col)

                    cell_text = ""
                    for cell_rel in cell.get("Relationships", []):
                        if cell_rel["Type"] == "CHILD":
                            for word_id in cell_rel["Ids"]:
                                word = block_map.get(word_id)
                                if word and word["BlockType"] == "WORD":
                                    cell_text += word.get("Text", "") + " "
                    cells[(row, col)] = cell_text.strip()

            if max_row > 0 and max_col > 0:
                rows = [
                    [cells.get((r, c), "") for c in range(1, max_col + 1)]
                    for r in range(1, max_row + 1)
                ]
                tables.append({"header": rows[0] if rows else [], "rows": rows})

    return {"text": "\n".join(lines), "lines": lines, "tables": tables}


# ---------------------------------------------------------------------------
# Backend implementations
# ---------------------------------------------------------------------------

def _textract_ocr_bytes(image_bytes: bytes, analyze_tables: bool = False) -> Dict:
    """
    Call AWS Textract.
    - analyze_tables=False  → detect_document_text  ($0.0015/page)
    - analyze_tables=True   → analyze_document + TABLES feature ($0.015/page)

    Returns {'text', 'lines', 'tables'}.
    """
    if textract_client is None:
        raise RuntimeError("Textract client not initialized. Call init_textract() first.")

    if analyze_tables:
        response = textract_client.analyze_document(
            Document={"Bytes": image_bytes},
            FeatureTypes=["TABLES"],
        )
    else:
        response = textract_client.detect_document_text(
            Document={"Bytes": image_bytes},
        )

    return _parse_textract_blocks(response.get("Blocks", []))


def _google_vision_ocr_bytes(image_bytes: bytes) -> Dict:
    """
    Return {'text', 'lines', 'tables': []} using Google Vision document_text_detection.
    """
    if google_client is None:
        raise RuntimeError("Google Vision client not initialized. Call init_google_vision() first.")

    from google.cloud import vision

    image = vision.Image(content=image_bytes)
    response = google_client.document_text_detection(image=image)
    if response.error.message:
        raise RuntimeError(response.error.message)

    full_text = response.full_text_annotation.text if response.full_text_annotation else ""
    lines: List[str] = []
    if response.full_text_annotation:
        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    words = ["".join(s.text for s in w.symbols) for w in paragraph.words]
                    if words:
                        lines.append(" ".join(words))

    return {"text": full_text, "lines": lines, "tables": []}


def _tesseract_ocr_bytes(image_bytes: bytes) -> Dict:
    """Return {'text', 'lines', 'tables': []} using local Tesseract."""
    try:
        import pytesseract
    except ImportError as e:
        raise RuntimeError("pytesseract is required for Tesseract OCR") from e

    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    text = pytesseract.image_to_string(img)
    lines = [ln for ln in text.splitlines() if ln.strip()]
    return {"text": text, "lines": lines, "tables": []}


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def ocr_image(image_bytes: bytes, backend: str = "google_vision") -> Dict:
    """
    Generic OCR wrapper. Returns dict with 'text', 'lines', and 'tables'.

    backend options
    ---------------
    'google_vision'  Google Cloud Vision — text only, cheapest for plain text.
    'textract'       AWS Textract — auto-detects tables via image heuristic;
                     uses analyze_document (TABLES) when a table is found,
                     detect_document_text otherwise.
    'auto'           Hybrid: Textract for table images, Google Vision for plain
                     text, Tesseract as final offline fallback.
    'tesseract'      Local Tesseract only — no cloud cost, lower accuracy.
    """
    if backend == "google_vision":
        try:
            return _google_vision_ocr_bytes(image_bytes)
        except Exception as primary_err:
            try:
                return _tesseract_ocr_bytes(image_bytes)
            except Exception:
                raise primary_err

    elif backend == "textract":
        has_table = _has_table_heuristic(image_bytes)
        try:
            return _textract_ocr_bytes(image_bytes, analyze_tables=has_table)
        except Exception as primary_err:
            try:
                return _tesseract_ocr_bytes(image_bytes)
            except Exception:
                raise primary_err

    elif backend == "auto":
        has_table = _has_table_heuristic(image_bytes)
        # Table image → Textract (best table structure recovery)
        if has_table and textract_client is not None:
            try:
                return _textract_ocr_bytes(image_bytes, analyze_tables=True)
            except Exception:
                pass  # fall through to Google Vision
        # Plain text image → Google Vision (cheaper, already integrated)
        if google_client is not None:
            try:
                return _google_vision_ocr_bytes(image_bytes)
            except Exception:
                pass  # fall through to Tesseract
        # Offline fallback
        return _tesseract_ocr_bytes(image_bytes)

    elif backend == "tesseract":
        return _tesseract_ocr_bytes(image_bytes)

    else:
        raise ValueError(f"Unknown OCR backend: {backend!r}. "
                         "Choose from: 'google_vision', 'textract', 'auto', 'tesseract'.")
