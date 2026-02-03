# app.py
import io
import json
import os
from typing import List, Dict, Optional

import streamlit as st
from pptx import Presentation

import pptx_reader
import preprocessor
import ocr_backend
import paginator
import template_filler
import utils

# Constants matching your requirements
TITLE_FONT_NAME = "Poppins"
BODY_FONT_NAME = "Poppins"
TITLE_PX = 20
BODY_PX = 12
DEFAULT_DPI = 300
DEFAULT_CONTINUATION_SUFFIX = "(CONTD...)"

st.set_page_config(page_title="PPT Normalizer — Phase B (OCR)", layout="wide")

st.title("PPT Normalizer — Phase B (MVP + OCR)")
st.markdown(
    """
Upload an input PPTX and optionally upload a standard template. The app will:
- Extract Title + Body from slides (prefer pptx text shapes).
- If body is missing and 'Use OCR on images' is enabled, OCR embedded images (Google Vision by default).
- Provide a per-slide editable preview.
- Paginate using paragraph heuristics or precise font-based pagination if you upload Poppins TTF.
- Fill your runtime template (preserving background) and produce a standardized PPTX you can download.
"""
)

# ---- Sidebar controls ----
with st.sidebar:
    st.header("Processing Options")
    use_ocr = st.checkbox("Use OCR on images (may be slower / cost money)", value=True)
    ocr_backend_choice = st.selectbox("OCR backend", ["google_vision", "tesseract"], index=0)
    ocr_scope = st.selectbox(
        "OCR scope",
        ["Only when text-shapes missing/ambiguous", "Always (process all slides)"],
        index=0,
    )
    dpi = st.number_input("Rasterization DPI (for measurement)", value=DEFAULT_DPI, min_value=72, max_value=600)
    continuation_style = st.text_input("Continuation suffix (default)", value=DEFAULT_CONTINUATION_SUFFIX)
    paragraph_chars_per_page = st.number_input("Chars per page (heuristic pagination)", value=1100, min_value=200)

    st.markdown("---")
    # Determine Google Vision status
google_vision_status = "INACTIVE"
status_icon = "🔴"  # Initialize as red dot by default

if st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
    try:
        ocr_backend.init_google_vision(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])
        google_vision_status = "ACTIVE"
        status_icon = "🟢"  # Change to green dot if Google Vision initializes successfully
    except Exception as e:
        st.warning("Google Vision initialization failed. Check your service account credentials.")

with st.sidebar:
    st.header("Processing Options")
    st.checkbox("Use OCR on images (may be slower / cost money)", value=True)
    ocr_backend_choice = st.selectbox("OCR backend", ["google_vision", "tesseract"], index=0)
    ocr_scope = st.selectbox(
        "OCR scope",
        ["Only when text-shapes missing/ambiguous", "Always (process all slides)"],
        index=0,
    )
    dpi = st.number_input("Rasterization DPI (for measurement)", value=300, min_value=72, max_value=600)
    continuation_style = st.text_input("Continuation suffix (default)", value="(CONTD...)")
    paragraph_chars_per_page = st.number_input("Chars per page (heuristic pagination)", value=1100, min_value=200)

    st.markdown("---")
    st.header("Google Vision Status")
    st.markdown(f"{status_icon} **Google Vision Status: {google_vision_status}**")

# ---- Main UI: Uploaders ----
col1, col2 = st.columns(2)
with col1:
    input_ppt = st.file_uploader("Upload input .pptx", type=["pptx"])
with col2:
    template_ppt = st.file_uploader("Upload template .pptx (optional)", type=["pptx"])

poppins_ttf = st.file_uploader("Upload Poppins .ttf (optional — improves pagination)", type=["ttf", "otf"])

# Persist session state containers
if "slides_meta" not in st.session_state:
    st.session_state["slides_meta"] = []
if "titles" not in st.session_state:
    st.session_state["titles"] = {}
if "bodies" not in st.session_state:
    st.session_state["bodies"] = {}

# Setup Google Vision client if requested
google_credentials_json = None
if uploaded_gcs_json is not None:
    google_credentials_json = uploaded_gcs_json.read().decode("utf-8")
else:
    # Check Streamlit secrets
    if st.secrets and "GOOGLE_SERVICE_ACCOUNT_JSON" in st.secrets:
        google_credentials_json = st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"]

if ocr_backend_choice == "google_vision":
    try:
        ocr_backend.init_google_vision(google_credentials_json)
        st.sidebar.success("Google Vision initialized (if credentials provided).")
    except Exception as e:
        st.sidebar.warning("Google Vision init failed or no credentials provided; will fallback to Tesseract if selected.")
        st.sidebar.write(e)


def analyze_and_preview():
    """
    Extract text shapes from pptx. If allowed and needed, run OCR on embedded images.
    Populate session_state['slides_meta'] and session_state['titles']/['bodies'].
    """
    if input_ppt is None:
        st.warning("Please upload an input .pptx file first.")
        return

    input_bytes = input_ppt.read()
    slides_meta = pptx_reader.extract_text_shapes(input_bytes)

    # For each slide, determine title/body; if missing body and OCR allowed, run OCR on image shapes
    for meta in slides_meta:
        slide_idx = meta["slide_index"]
        title = ""
        body = ""

        # Prefer placeholder title, else first shape line
        if meta.get("title_text"):
            title = meta["title_text"]
        else:
            # If there are any text shapes, take first non-empty line as title
            if meta.get("text_shapes"):
                first_text = meta["text_shapes"][0]["text"].strip()
                first_line = first_text.splitlines()[0].strip() if first_text else ""
                title = first_line or f"Slide {slide_idx + 1}"
                # Remaining lines go to body if present
                remainder = "\n".join(first_text.splitlines()[1:]).strip()
                if remainder:
                    body = remainder

        # Compose body from text shapes (excluding the title shape)
        if meta.get("text_shapes"):
            # Combine all text shapes except the one used as title
            body_parts = []
            for s in meta["text_shapes"]:
                txt = s["text"].strip()
                if not txt:
                    continue
                # Skip if exact match with title
                if title and txt == title:
                    continue
                body_parts.append(txt)
            if body_parts:
                body = "\n\n".join(body_parts).strip()

        # Decide if OCR is needed
        needs_ocr = False
        if ocr_scope.startswith("Only") and (not body or len(body) < 20):
            needs_ocr = True
        elif ocr_scope.startswith("Always"):
            needs_ocr = True

        ocr_text = ""
        if use_ocr and needs_ocr:
            # OCR embedded image shapes first
            image_shapes = meta.get("image_shapes", [])
            ocr_texts = []
            for img_meta in image_shapes:
                img_bytes = img_meta.get("image_bytes")
                if not img_bytes:
                    continue
                # preprocess
                try:
                    img_bytes_proc = preprocessor.preprocess_image(img_bytes, dpi=dpi)
                except Exception:
                    img_bytes_proc = img_bytes
                try:
                    ocr_result = ocr_backend.ocr_image(img_bytes_proc, backend=ocr_backend_choice)
                    if ocr_result and ocr_result.get("text"):
                        ocr_texts.append(ocr_result["text"].strip())
                except Exception as e:
                    st.warning(f"OCR failed on a picture on slide {slide_idx+1}: {e}")
            if ocr_texts:
                ocr_text = "\n\n".join(ocr_texts)

            # NOTE: full-slide rasterization is not implemented here due to environment constraints.
            # If you want full-slide OCR, see README for instructions to enable slide->image conversion via LibreOffice or other tools.

        # Merge OCR text into body if body empty or short
        if (not body or len(body) < 20) and ocr_text:
            body = (body + "\n\n" + ocr_text).strip() if body else ocr_text

        # If still no body, set empty string
        if not body:
            body = ""

        st.session_state["titles"][slide_idx] = title
        st.session_state["bodies"][slide_idx] = body

    st.session_state["slides_meta"] = slides_meta
    st.success(f"Analyzed {len(slides_meta)} slides. Review & edit below.")


def render_preview_and_edit():
    """
    Show per-slide preview with editable fields and allow per-slide OCR re-run on images.
    """
    slides_meta = st.session_state.get("slides_meta", [])
    if not slides_meta:
        st.info("No analysis available yet. Click 'Analyze & Preview' after uploading input PPTX.")
        return

    for meta in slides_meta:
        idx = meta["slide_index"]
        header = f"Slide {idx + 1}"
        with st.expander(header, expanded=False):
            title_key = f"title_{idx}"
            body_key = f"body_{idx}"
            current_title = st.session_state["titles"].get(idx, "")
            current_body = st.session_state["bodies"].get(idx, "")

            new_title = st.text_input(f"Title (Slide {idx+1})", value=current_title, key=title_key)
            new_body = st.text_area(f"Body (Slide {idx+1})", value=current_body, height=200, key=body_key)

            # Buttons for per-slide OCR re-run (images)
            colA, colB = st.columns([1, 3])
            with colA:
                if st.button("Re-run OCR on images for this slide", key=f"ocr_rerun_{idx}"):
                    img_shapes = meta.get("image_shapes", [])
                    ocr_texts = []
                    for img_meta in img_shapes:
                        img_bytes = img_meta.get("image_bytes")
                        if not img_bytes:
                            continue
                        try:
                            img_bytes_proc = preprocessor.preprocess_image(img_bytes, dpi=dpi)
                        except Exception:
                            img_bytes_proc = img_bytes
                        try:
                            res = ocr_backend.ocr_image(img_bytes_proc, backend=ocr_backend_choice)
                            if res and res.get("text"):
                                ocr_texts.append(res["text"].strip())
                        except Exception as e:
                            st.warning(f"OCR failed on slide {idx+1} image: {e}")
                    if ocr_texts:
                        new_body = (new_body + "\n\n" + "\n\n".join(ocr_texts)).strip() if new_body else "\n\n".join(ocr_texts)
                        st.session_state["bodies"][idx] = new_body
                        st.success("OCR results appended to body for this slide.")
                    else:
                        st.info("No OCR text detected on images for this slide.")

            with colB:
                st.write("Preview of extracted content. Edit as needed before final generation.")

            # Save edits back to session_state
            st.session_state["titles"][idx] = new_title
            st.session_state["bodies"][idx] = new_body


def generate_and_download():
    """
    Build pages based on session_state titles & bodies, paginate, fill template (if any), and produce final pptx bytes.
    """
    slides_meta = st.session_state.get("slides_meta", [])
    if not slides_meta:
        st.warning("No slides to generate from. Run Analyze & Preview first.")
        return

    # Build pages list: for each slide, create 1..N pages
    pages = []
    for meta in slides_meta:
        idx = meta["slide_index"]
        title = st.session_state["titles"].get(idx, f"Slide {idx+1}")
        body = st.session_state["bodies"].get(idx, "")

        # Paginate using font metrics if TTF provided, else paragraph heuristic
        if poppins_ttf is not None:
            # Try precise pagination
            try:
                font_bytes = poppins_ttf.read()
                pages_texts = paginator.split_by_font_metrics(
                    body,
                    font_bytes=font_bytes,
                    box_width_px=None,  # will be inferred from template if possible
                    box_height_px=None,
                    font_size_px=BODY_PX,
                    dpi=dpi,
                )
            except Exception as e:
                st.warning(f"Precise pagination failed, falling back to heuristic: {e}")
                pages_texts = paginator.split_by_paragraphs(body, chars_per_page=paragraph_chars_per_page)
        else:
            pages_texts = paginator.split_by_paragraphs(body, chars_per_page=paragraph_chars_per_page)

        for i, page_text in enumerate(pages_texts):
            page_title = title if i == 0 else f"{title} {continuation_style}"
            pages.append({"title": page_title, "body": page_text})

    # Fill template or use default
    input_template_bytes = template_ppt.read() if template_ppt is not None else None
    try:
        out_bytes = template_filler.fill_template_with_pages(
            template_bytes=input_template_bytes,
            pages=pages,
            title_font=TITLE_FONT_NAME,
            body_font=BODY_FONT_NAME,
            title_font_pt=None,
            body_font_pt=None,
        )
    except Exception as e:
        st.error(f"Failed to generate PPTX: {e}")
        return

    st.success("Generated standardized PPTX.")
    st.download_button(
        label="Download standardized PPTX",
        data=out_bytes,
        file_name=f"standardized_{input_ppt.name if input_ppt is not None else 'output'}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

# ---- Buttons ----
colA, colB, colC = st.columns([1, 1, 1])
with colA:
    if st.button("Analyze & Preview"):
        analyze_and_preview()
with colB:
    if st.button("Render Preview"):
        render_preview_and_edit()
with colC:
    if st.button("Generate & Download"):
        generate_and_download()

st.markdown("---")
st.markdown(
    """
Notes:
- OCR currently operates on embedded images inside slides (common case for screenshots). Full-slide rasterization is not provided by default since converting PPTX slides to images often requires LibreOffice or other external tools on the host. If you need full-slide OCR, I can add optional conversion steps (requires LibreOffice or Windows PowerPoint automation) — tell me and I will add it.
- For Google Vision, please add your service account JSON to Streamlit Secrets (key: GOOGLE_SERVICE_ACCOUNT_JSON) or upload it in the sidebar for a session.
"""
)
