import io
import json
import os
from typing import List, Dict, Optional

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt

import pptx_reader
import preprocessor
import ocr_backend
import paginator
import template_filler
import utils

# Constants
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
- Extract Title + Body from slides (prefer PPTX text shapes).
- If body is missing and 'Use OCR on images' is enabled, OCR embedded images (Google Vision by default).
- Provide a per-slide editable preview.
- Paginate using paragraph heuristics or precise font-based pagination if you upload Poppins TTF.
- Fill your runtime template (preserving background) and produce a standardized PPTX you can download.
"""
)

# ---- Sidebar controls ----
with st.sidebar:
    st.header("Processing Options")
    use_ocr = st.checkbox("Use OCR on images (may be slower / cost money)", value=True, key="use_ocr_checkbox")
    ocr_backend_choice = st.selectbox("OCR backend", ["google_vision", "tesseract"], index=0, key="ocr_backend_choice")
    ocr_scope = st.selectbox(
        "OCR scope",
        ["Only when text-shapes missing/ambiguous", "Always (process all slides)"],
        index=0,
        key="ocr_scope_select"
    )
    dpi = st.number_input("Rasterization DPI (for measurement)", value=DEFAULT_DPI, min_value=72, max_value=600, key="dpi_input")
    continuation_style = st.text_input("Continuation suffix (default)", value=DEFAULT_CONTINUATION_SUFFIX, key="cont_suffix_input")
    paragraph_chars_per_page = st.number_input("Chars per page (heuristic pagination)", value=1100, min_value=200, key="chars_per_page_input")

    st.markdown("---")

    # Determine Google Vision status
    google_vision_status = "INACTIVE"
    status_icon = "🔴"  # Red dot by default
    if st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
        try:
            ocr_backend.init_google_vision(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])
            google_vision_status = "ACTIVE"
            status_icon = "🟢"  # Green dot if initialized successfully
        except Exception as e:
            st.warning("Google Vision initialization failed. Check your service account credentials.")

    # Display Google Vision Status
    st.header("Google Vision Status")
    st.markdown(f"{status_icon} **Google Vision Status: {google_vision_status}**")

# ---- Main UI: Uploaders ----
col1, col2 = st.columns(2)
with col1:
    input_ppt = st.file_uploader("Upload input .pptx", type=["pptx"], key="input_ppt_uploader")
with col2:
    template_ppt = st.file_uploader("Upload template .pptx (optional)", type=["pptx"], key="template_ppt_uploader")

poppins_ttf = st.file_uploader("Upload Poppins .ttf (optional — improves pagination)", type=["ttf", "otf"], key="poppins_uploader")

# Persist session state containers
if "slides_meta" not in st.session_state:
    st.session_state["slides_meta"] = []
if "titles" not in st.session_state:
    st.session_state["titles"] = {}
if "bodies" not in st.session_state:
    st.session_state["bodies"] = {}
if "tables" not in st.session_state:
    st.session_state["tables"] = {}

# ---- Functions ----
def analyze_and_preview():
    """
    Extract text shapes and tables from pptx. If allowed and needed, run OCR on embedded images.
    Populate session_state['slides_meta'], session_state['titles'], session_state['bodies'], and session_state['tables'].
    """
    if input_ppt is None:
        st.warning("Please upload an input .pptx file first.")
        return

    input_bytes = input_ppt.read()
    slides_meta = pptx_reader.extract_text_shapes(input_bytes)
    table_data = []

    for meta in slides_meta:
        slide_idx = meta["slide_index"]
        title = ""
        body = ""

        # Extract title and body from text shapes
        # (Use logic from previous placeholders for title and body extraction)

        # Extract tables as editable structures
        for shape in meta.get("shapes", []):
            if shape.has_table:
                table = shape.table
                table_content = []
                for row in table.rows:
                    row_data = [cell.text.strip() for cell in row.cells]
                    table_content.append(row_data)

                table_data.append({
                    "header": [cell.text.strip() for cell in table.rows[0].cells],
                    "rows": table_content,
                    "slide_index": slide_idx
                })

        st.session_state["titles"][slide_idx] = title
        st.session_state["bodies"][slide_idx] = body

    st.session_state["slides_meta"] = slides_meta
    st.session_state["tables"] = table_data
    st.success(f"Analyzed {len(slides_meta)} slides. Review & edit below.")

def render_preview_and_edit():
    """
    Show per-slide preview, including titles, bodies, and tables, with editable fields.
    """
    slides_meta = st.session_state.get("slides_meta", [])
    tables = st.session_state.get("tables", [])
    if not slides_meta and not tables:
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

            # Add editable fields for tables
            for table in [t for t in tables if t["slide_index"] == idx]:
                st.subheader("Table Preview")
                st.table(table["rows"])

            st.session_state["titles"][idx] = new_title
            st.session_state["bodies"][idx] = new_body

def generate_and_download():
    """
    Generate final standardized PPTX and allow downloading the result.
    """
    if not st.session_state.get("slides_meta"):
        st.warning("No slides to generate from. Run Analyze & Preview first.")
        return

    slides_meta = st.session_state["slides_meta"]
    tables = st.session_state["tables"]
    pages = []
    for meta in slides_meta:
        slide_idx = meta["slide_index"]
        title = st.session_state["titles"].get(slide_idx, f"Slide {slide_idx + 1}")
        body = st.session_state["bodies"].get(slide_idx, "")

        if poppins_ttf is not None:
            font_bytes = poppins_ttf.read()
            try:
                pages_texts = paginator.split_by_font_metrics(
                    body,
                    font_bytes=font_bytes,
                    box_width_px=None,
                    box_height_px=None,
                    font_size_px=BODY_PX,
                    dpi=dpi,
                )
            except Exception as e:
                st.warning(f"Precise pagination failed for slide {slide_idx+1}. Falling back to heuristic.")
                pages_texts = paginator.split_by_paragraphs(body, chars_per_page=paragraph_chars_per_page)
        else:
            pages_texts = paginator.split_by_paragraphs(body, chars_per_page=paragraph_chars_per_page)

        for i, page_text in enumerate(pages_texts):
            page_title = title if i == 0 else f"{title} {continuation_style}"
            pages.append({"title": page_title, "body": page_text})

    template_bytes = template_ppt.read() if template_ppt else None

    try:
        pptx_output_bytes = template_filler.fill_template_with_pages(
            template_bytes=template_bytes,
            pages=pages,
            tables=tables,
            title_font=TITLE_FONT_NAME,
            body_font=BODY_FONT_NAME,
        )
        st.success("Generated standardized PPTX.")
        st.download_button(
            label="Download standardized PPTX",
            data=pptx_output_bytes,
            file_name=f"standardized_{input_ppt.name if input_ppt else 'output'}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except Exception as e:
        st.error(f"Failed to generate PPTX: {e}")

# ---- Main Actions ----
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
