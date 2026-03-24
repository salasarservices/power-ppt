import io
import traceback as tb
from typing import Dict, List, Optional

import streamlit as st

import ocr_backend
import paginator
import preprocessor
import pptx_reader
import template_filler
import utils

# ── Constants ──────────────────────────────────────────────────────────────────
TITLE_FONT_NAME = "Poppins"
BODY_FONT_NAME  = "Poppins"
DEFAULT_DPI = 300
DEFAULT_CONTINUATION_SUFFIX = "(CONTD...)"
LOGO_URL = "https://ik.imagekit.io/salasarservices/Salasar-Logo-new.png"

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PowerPPT — Salasar Services",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

:root {
  --primary:     #0D2B5E;
  --primary-lt:  #1A3C6E;
  --primary-dim: rgba(13,43,94,0.08);
  --accent:      #E05A1E;
  --bg:          #EEF1F7;
  --card:        #FFFFFF;
  --border:      #DDE3EE;
  --text:        #1A202C;
  --muted:       #64748B;
  --success:     #10B981;
  --warning:     #F59E0B;
  --danger:      #EF4444;
  --radius:      10px;
  --shadow:      0 2px 12px rgba(13,43,94,.09);
}

/* ── Base ── */
html, body, [class*="css"] { font-family: 'Poppins', sans-serif !important; }
.main .block-container {
  padding: 1.5rem 2rem 3rem;
  max-width: 1500px;
  background: var(--bg);
}
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: var(--primary) !important;
  border-right: none !important;
}
[data-testid="stSidebar"] > div:first-child {
  padding-top: 0 !important;
}
[data-testid="stSidebar"] * { color: #CBD5E1 !important; }

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
  color: #94A3B8 !important;
  font-size: 0.65rem !important;
  font-weight: 700 !important;
  letter-spacing: 0.12em !important;
  text-transform: uppercase !important;
  margin: 1.4rem 0 0.5rem !important;
}
[data-testid="stSidebar"] label {
  color: #94A3B8 !important;
  font-size: 0.75rem !important;
  font-weight: 500 !important;
}
[data-testid="stSidebar"] .stSelectbox > div > div,
[data-testid="stSidebar"] input {
  background: rgba(255,255,255,0.07) !important;
  border: 1px solid rgba(255,255,255,0.13) !important;
  color: #E2E8F0 !important;
  border-radius: 7px !important;
  font-size: 0.78rem !important;
}
[data-testid="stSidebar"] hr {
  border-color: rgba(255,255,255,0.1) !important;
  margin: 1rem 0 !important;
}
[data-testid="stSidebar"] .stCheckbox span { color: #CBD5E1 !important; font-size: 0.78rem !important; }

/* ── Topbar ── */
.ss-topbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 0.75rem 1.5rem;
  margin-bottom: 1.4rem;
  box-shadow: var(--shadow);
}
.ss-topbar-left { display: flex; align-items: center; gap: 14px; }
.ss-topbar-logo img { height: 38px; object-fit: contain; }
.ss-topbar-divider {
  width: 1px; height: 34px;
  background: var(--border);
}
.ss-topbar-title {
  font-size: 1rem; font-weight: 700;
  color: var(--primary); line-height: 1.2;
}
.ss-topbar-subtitle {
  font-size: 0.68rem; color: var(--muted);
  margin-top: 2px;
}
.ss-topbar-right {
  display: flex; align-items: center; gap: 8px;
}
.ss-topbar-pill {
  background: var(--primary-dim);
  color: var(--primary) !important;
  font-size: 0.68rem; font-weight: 600;
  padding: 3px 10px; border-radius: 20px;
  letter-spacing: 0.04em;
}

/* ── Step bar ── */
.ss-steps {
  display: flex; align-items: center;
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 0.85rem 1.5rem;
  margin-bottom: 1.4rem;
  box-shadow: var(--shadow);
  gap: 0;
}
.ss-step { display: flex; align-items: center; flex: 1; }
.ss-step-circle {
  width: 26px; height: 26px; border-radius: 50%;
  background: var(--border); color: var(--muted);
  font-size: 0.7rem; font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}
.ss-step-circle.active { background: var(--primary); color: #fff; }
.ss-step-circle.done   { background: var(--success); color: #fff; }
.ss-step-body { margin-left: 9px; }
.ss-step-name {
  font-size: 0.72rem; font-weight: 700;
  color: var(--muted); line-height: 1.2;
}
.ss-step-name.active { color: var(--primary); }
.ss-step-name.done   { color: var(--success); }
.ss-step-desc { font-size: 0.63rem; color: var(--muted); }
.ss-step-line { flex: 1; height: 2px; background: var(--border); margin: 0 12px; }
.ss-step-line.done { background: var(--success); }

/* ── Cards ── */
.ss-card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 1.25rem 1.5rem;
  box-shadow: var(--shadow);
  margin-bottom: 1rem;
}
.ss-card-header {
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 0.85rem;
}
.ss-card-title {
  font-size: 0.72rem; font-weight: 700;
  letter-spacing: 0.1em; text-transform: uppercase;
  color: var(--muted);
}
.ss-card-icon {
  font-size: 1rem; opacity: 0.6;
}

/* ── Metrics ── */
.ss-metrics {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 1rem;
  margin-bottom: 1.4rem;
}
.ss-metric {
  background: var(--card);
  border: 1px solid var(--border);
  border-top: 3px solid var(--primary);
  border-radius: var(--radius);
  padding: 1rem 1.25rem;
  box-shadow: var(--shadow);
}
.ss-metric-value {
  font-size: 1.8rem; font-weight: 700;
  color: var(--primary); line-height: 1;
}
.ss-metric-label {
  font-size: 0.66rem; font-weight: 600;
  color: var(--muted); text-transform: uppercase;
  letter-spacing: 0.08em; margin-top: 5px;
}

/* ── Status badge ── */
.ss-badge {
  display: inline-flex; align-items: center;
  gap: 5px; padding: 3px 9px;
  border-radius: 20px; font-size: 0.68rem; font-weight: 600;
}
.ss-badge-on  { background: #D1FAE5; color: #065F46; }
.ss-badge-off { background: #FEE2E2; color: #991B1B; }
.ss-badge-dot { width: 6px; height: 6px; border-radius: 50%; }
.ss-badge-on .ss-badge-dot  { background: #10B981; }
.ss-badge-off .ss-badge-dot { background: #EF4444; }

/* ── Service row ── */
.ss-svc-row {
  display: flex; align-items: center;
  justify-content: space-between;
  padding: 8px 0;
  border-bottom: 1px solid rgba(255,255,255,0.08);
}
.ss-svc-row:last-child { border-bottom: none; }
.ss-svc-name  { font-size: 0.75rem; font-weight: 600; color: #E2E8F0 !important; }
.ss-svc-desc  { font-size: 0.65rem; color: #64748B !important; margin-top: 1px; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 4px; gap: 2px;
  box-shadow: var(--shadow);
  margin-bottom: 1rem;
}
.stTabs [data-baseweb="tab"] {
  border-radius: 7px !important;
  font-size: 0.78rem !important; font-weight: 600 !important;
  color: var(--muted) !important;
  padding: 0.45rem 1.3rem !important;
  background: transparent !important;
  border: none !important;
}
.stTabs [aria-selected="true"] {
  background: var(--primary) !important;
  color: #fff !important;
}
.stTabs [data-baseweb="tab-panel"] { padding: 0 !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
  background: var(--card) !important;
  border: 2px dashed var(--border) !important;
  border-radius: var(--radius) !important;
}
[data-testid="stFileUploader"]:hover { border-color: var(--primary) !important; }
[data-testid="stFileUploader"] small { font-size: 0.7rem !important; }

/* ── Buttons ── */
.stButton > button {
  border-radius: 8px !important;
  font-family: 'Poppins', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  padding: 0.55rem 1.4rem !important;
  border: none !important;
  transition: all 0.18s ease !important;
  background: var(--primary) !important;
  color: #fff !important;
  width: 100%;
}
.stButton > button:hover {
  background: var(--primary-lt) !important;
  transform: translateY(-1px);
  box-shadow: 0 4px 14px rgba(13,43,94,0.25) !important;
}
.stDownloadButton > button {
  background: var(--success) !important;
  color: #fff !important;
  border-radius: 8px !important;
  font-family: 'Poppins', sans-serif !important;
  font-weight: 600 !important; font-size: 0.85rem !important;
  padding: 0.65rem 2rem !important;
  border: none !important; width: 100%;
  transition: all 0.18s ease !important;
}
.stDownloadButton > button:hover {
  background: #059669 !important;
  transform: translateY(-1px);
  box-shadow: 0 4px 14px rgba(16,185,129,0.3) !important;
}

/* ── Expanders (slide editor cards) ── */
[data-testid="stExpander"] {
  background: var(--card) !important;
  border: 1px solid var(--border) !important;
  border-radius: var(--radius) !important;
  box-shadow: var(--shadow);
  margin-bottom: 0.6rem !important;
  overflow: hidden;
}
[data-testid="stExpander"] summary {
  font-size: 0.82rem !important; font-weight: 600 !important;
  color: var(--primary) !important;
  padding: 0.7rem 1rem !important;
}
[data-testid="stExpander"] summary:hover { background: var(--primary-dim) !important; }

/* ── Inputs ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
  border-radius: 7px !important;
  border: 1px solid var(--border) !important;
  font-family: 'Poppins', sans-serif !important;
  font-size: 0.8rem !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
  border-color: var(--primary) !important;
  box-shadow: 0 0 0 2px var(--primary-dim) !important;
}

/* ── Alerts ── */
.stAlert { border-radius: var(--radius) !important; font-size: 0.8rem !important; }

/* ── Upload hint label ── */
.ss-upload-hint {
  font-size: 0.7rem; color: var(--muted);
  margin-bottom: 4px; display: block;
}
.ss-upload-required {
  font-size: 0.65rem; font-weight: 600;
  color: var(--accent); margin-left: 4px;
}

/* ── Section heading ── */
.ss-section {
  font-size: 0.78rem; font-weight: 700;
  color: var(--primary); letter-spacing: 0.02em;
  border-left: 3px solid var(--primary);
  padding-left: 10px; margin-bottom: 0.85rem;
}

/* ── Empty state ── */
.ss-empty {
  text-align: center; padding: 3rem 1rem;
  color: var(--muted);
}
.ss-empty-icon { font-size: 2.5rem; margin-bottom: 0.75rem; }
.ss-empty-title { font-size: 0.9rem; font-weight: 600; color: var(--text); }
.ss-empty-desc  { font-size: 0.75rem; margin-top: 4px; }

/* ── Sidebar logo block ── */
.ss-sidebar-brand {
  background: rgba(0,0,0,0.2);
  padding: 1.4rem 1rem 1rem;
  text-align: center;
  margin-bottom: 0.25rem;
}
.ss-sidebar-brand img {
  max-width: 150px; height: auto;
  filter: brightness(0) invert(1);
  opacity: 0.92;
}
.ss-sidebar-appname {
  font-size: 0.6rem !important; font-weight: 700 !important;
  letter-spacing: 0.18em !important; text-transform: uppercase !important;
  color: rgba(255,255,255,0.35) !important;
  margin-top: 8px !important;
}
</style>
""", unsafe_allow_html=True)


# ── Session state init ─────────────────────────────────────────────────────────
for key, default in [
    ("slides_meta", []),
    ("titles", {}),
    ("bodies", {}),
    ("tables", []),
    ("pptx_output_bytes", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ── Cached client initializers ─────────────────────────────────────────────────
@st.cache_resource
def _init_google_vision_cached(credentials_json: str):
    ocr_backend.init_google_vision(credentials_json)

@st.cache_resource
def _init_textract_cached(key_id: str, secret_key: str, region: str):
    ocr_backend.init_textract(key_id, secret_key, region)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    # Brand block
    st.markdown(f"""
    <div class="ss-sidebar-brand">
        <img src="{LOGO_URL}" alt="Salasar Services">
        <div class="ss-sidebar-appname">PowerPPT Normalizer</div>
    </div>
    """, unsafe_allow_html=True)

    # OCR settings
    st.markdown("### OCR Settings")
    use_ocr = st.checkbox(
        "Enable OCR on image slides",
        value=True,
        key="use_ocr_checkbox",
    )
    ocr_backend_choice = st.selectbox(
        "OCR Engine",
        ["auto", "textract", "google_vision", "tesseract"],
        index=0,
        key="ocr_backend_choice",
        help=(
            "auto — smart routing: Textract for tables, Google Vision for text\n"
            "textract — AWS Textract (superior table extraction)\n"
            "google_vision — Google Cloud Vision (plain text)\n"
            "tesseract — local offline, no cloud cost"
        ),
    )
    ocr_scope = st.selectbox(
        "OCR Scope",
        ["Only when text shapes are missing", "Always (all slides)"],
        index=0,
        key="ocr_scope_select",
    )

    st.markdown("### Advanced")
    dpi = st.number_input(
        "Rasterization DPI", value=DEFAULT_DPI,
        min_value=72, max_value=600, key="dpi_input",
    )
    continuation_style = st.text_input(
        "Continuation suffix",
        value=DEFAULT_CONTINUATION_SUFFIX,
        key="cont_suffix_input",
    )
    paragraph_chars_per_page = st.number_input(
        "Chars per page (pagination)",
        value=1100, min_value=200, key="chars_per_page_input",
    )

    st.markdown("---")

    # Service status — initialise clients once, show live badges
    gv_active = False
    tx_active = False

    if st.secrets.get("GOOGLE_SERVICE_ACCOUNT_JSON"):
        try:
            _init_google_vision_cached(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])
            gv_active = True
        except Exception:
            pass

    if st.secrets.get("AWS_ACCESS_KEY_ID") and st.secrets.get("AWS_SECRET_ACCESS_KEY"):
        try:
            _init_textract_cached(
                st.secrets["AWS_ACCESS_KEY_ID"],
                st.secrets["AWS_SECRET_ACCESS_KEY"],
                st.secrets.get("AWS_DEFAULT_REGION", "us-east-1"),
            )
            tx_active = True
        except Exception:
            pass

    def _badge(active: bool) -> str:
        cls = "ss-badge-on" if active else "ss-badge-off"
        lbl = "Active" if active else "Inactive"
        return f'<span class="ss-badge {cls}"><span class="ss-badge-dot"></span>{lbl}</span>'

    st.markdown("### Service Status")
    st.markdown(f"""
    <div class="ss-svc-row">
        <div>
            <div class="ss-svc-name">Google Vision</div>
            <div class="ss-svc-desc">Plain-text OCR</div>
        </div>
        {_badge(gv_active)}
    </div>
    <div class="ss-svc-row">
        <div>
            <div class="ss-svc-name">AWS Textract</div>
            <div class="ss-svc-desc">Table-aware OCR</div>
        </div>
        {_badge(tx_active)}
    </div>
    <div class="ss-svc-row">
        <div>
            <div class="ss-svc-name">Tesseract</div>
            <div class="ss-svc-desc">Offline fallback</div>
        </div>
        {_badge(True)}
    </div>
    """, unsafe_allow_html=True)


# ── Topbar ─────────────────────────────────────────────────────────────────────
analyzed   = len(st.session_state["slides_meta"]) > 0
exportable = st.session_state["pptx_output_bytes"] is not None

st.markdown(f"""
<div class="ss-topbar">
    <div class="ss-topbar-left">
        <div class="ss-topbar-logo">
            <img src="{LOGO_URL}" alt="Salasar Services">
        </div>
        <div class="ss-topbar-divider"></div>
        <div>
            <div class="ss-topbar-title">PowerPPT Normalizer</div>
            <div class="ss-topbar-subtitle">Intelligent slide standardisation &amp; OCR pipeline</div>
        </div>
    </div>
    <div class="ss-topbar-right">
        <span class="ss-topbar-pill">Phase B — MVP</span>
        <span class="ss-topbar-pill">{'Analyzed: ' + str(len(st.session_state['slides_meta'])) + ' slides' if analyzed else 'No file analyzed'}</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ── Step indicator ─────────────────────────────────────────────────────────────
def _step(num: str, name: str, desc: str, state: str, line_done: bool = False) -> str:
    circle_cls = f"ss-step-circle {state}"
    label_cls  = f"ss-step-name {state}"
    check = "✓" if state == "done" else num
    line  = f'<div class="ss-step-line{"  done" if line_done else ""}"></div>' if num != "4" else ""
    return f"""
    <div class="ss-step">
        <div class="{circle_cls}">{check}</div>
        <div class="ss-step-body">
            <div class="{label_cls}">{name}</div>
            <div class="ss-step-desc">{desc}</div>
        </div>
    </div>
    {line}
    """

has_input  = "input_file_uploaded" in st.session_state and st.session_state["input_file_uploaded"]
step1 = "done" if analyzed else "active"
step2 = ("done" if exportable else "active") if analyzed else "pending"
step3 = "active" if exportable else ("active" if analyzed else "pending")
step4 = "done" if exportable else "pending"

st.markdown(f"""
<div class="ss-steps">
    {_step("1", "Upload", "Input PPTX &amp; template", step1, line_done=analyzed)}
    {_step("2", "Analyse", "Extract &amp; OCR content", step2, line_done=exportable)}
    {_step("3", "Review", "Edit per-slide content", step3, line_done=exportable)}
    {_step("4", "Export", "Generate &amp; download", step4, line_done=False)}
</div>
""", unsafe_allow_html=True)


# ── Metrics row (shown after analysis) ────────────────────────────────────────
if analyzed:
    slides_meta = st.session_state["slides_meta"]
    total_slides = len(slides_meta)
    slides_with_body  = sum(1 for m in slides_meta if st.session_state["bodies"].get(m["slide_index"], "").strip())
    slides_with_table = len({t["slide_index"] for t in st.session_state["tables"]})
    empty_slides = total_slides - slides_with_body

    st.markdown(f"""
    <div class="ss-metrics">
        <div class="ss-metric">
            <div class="ss-metric-value">{total_slides}</div>
            <div class="ss-metric-label">Total Slides</div>
        </div>
        <div class="ss-metric">
            <div class="ss-metric-value">{slides_with_body}</div>
            <div class="ss-metric-label">Slides with Text</div>
        </div>
        <div class="ss-metric">
            <div class="ss-metric-value">{slides_with_table}</div>
            <div class="ss-metric-label">Slides with Tables</div>
        </div>
        <div class="ss-metric" style="border-top-color: {'var(--warning)' if empty_slides > 0 else 'var(--success)'}">
            <div class="ss-metric-value" style="color: {'var(--warning)' if empty_slides > 0 else 'var(--success)'}">
                {empty_slides}
            </div>
            <div class="ss-metric-label">Empty / OCR-needed</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


# ── Main tabs ──────────────────────────────────────────────────────────────────
tab_upload, tab_review, tab_export = st.tabs([
    "  01 · Upload & Analyse  ",
    "  02 · Review & Edit  ",
    "  03 · Generate & Export  ",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Upload & Analyse
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown('<div class="ss-section">File Inputs</div>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2, gap="large")

    with col_a:
        st.markdown("""
        <div class="ss-card-title" style="margin-bottom:6px;">
            Input Presentation <span class="ss-upload-required">* required</span>
        </div>
        <div class="ss-upload-hint">The raw PPTX file you want to normalise.</div>
        """, unsafe_allow_html=True)
        input_ppt = st.file_uploader(
            "input_pptx", type=["pptx"], label_visibility="collapsed",
            key="input_ppt_uploader",
        )
        if input_ppt:
            st.success(f"Loaded: **{input_ppt.name}** ({input_ppt.size / 1024:.1f} KB)")

    with col_b:
        st.markdown("""
        <div class="ss-card-title" style="margin-bottom:6px;">
            Brand Template <span style="font-size:0.63rem;color:var(--muted);margin-left:4px;">optional</span>
        </div>
        <div class="ss-upload-hint">Your standard PPTX template (preserves backgrounds &amp; brand styles).</div>
        """, unsafe_allow_html=True)
        template_ppt = st.file_uploader(
            "template_pptx", type=["pptx"], label_visibility="collapsed",
            key="template_ppt_uploader",
        )
        if template_ppt:
            st.success(f"Loaded: **{template_ppt.name}** ({template_ppt.size / 1024:.1f} KB)")

    st.markdown('<div class="ss-section" style="margin-top:1.2rem;">Font (Optional)</div>', unsafe_allow_html=True)
    poppins_ttf = st.file_uploader(
        "Upload Poppins .ttf / .otf for precise font-metric pagination",
        type=["ttf", "otf"], key="poppins_uploader",
    )

    st.markdown('<div style="margin-top:1.5rem;"></div>', unsafe_allow_html=True)
    st.markdown('<div class="ss-section">Run Analysis</div>', unsafe_allow_html=True)

    col_btn, col_info = st.columns([1, 2], gap="large")
    with col_btn:
        run_analyse = st.button("Analyse Presentation", key="btn_analyse", use_container_width=True)
    with col_info:
        st.markdown("""
        <div style="padding:0.6rem 0; font-size:0.75rem; color:var(--muted); line-height:1.7;">
            Extracts titles, body text and tables from all slides.<br>
            If OCR is enabled, image-based slides are processed automatically.
        </div>
        """, unsafe_allow_html=True)

    if run_analyse:
        if input_ppt is None:
            st.error("Please upload an input PPTX file before running analysis.")
        else:
            with st.spinner("Analysing presentation…"):
                input_bytes = input_ppt.read()
                slides_meta = pptx_reader.extract_text_shapes(input_bytes)

                if not slides_meta:
                    st.error("No content could be extracted from this presentation.")
                else:
                    table_data: List[Dict] = []

                    for meta in slides_meta:
                        slide_idx = meta["slide_index"]
                        title = meta.get("title_text", f"Slide {slide_idx + 1}")
                        body  = meta.get("body_text", "")

                        # Native PPTX table shapes
                        for shape in meta.get("shapes", []):
                            if hasattr(shape, "has_table") and shape.has_table:
                                rows = [
                                    [cell.text.strip() for cell in row.cells]
                                    for row in shape.table.rows
                                ]
                                if rows:
                                    table_data.append({
                                        "header": rows[0],
                                        "rows": rows,
                                        "slide_index": slide_idx,
                                    })

                        # OCR fallback
                        always_ocr = st.session_state.get("ocr_scope_select", "") == "Always (all slides)"
                        if use_ocr and (not body.strip() or always_ocr):
                            for img_shape in meta.get("image_shapes", []):
                                try:
                                    processed = preprocessor.preprocess_image(img_shape["image_bytes"])
                                    result = ocr_backend.ocr_image(processed, backend=ocr_backend_choice)
                                    if result.get("text", "").strip():
                                        body = (body + "\n\n" + result["text"]).strip() if body else result["text"].strip()
                                    for tbl in result.get("tables", []):
                                        if tbl.get("rows"):
                                            table_data.append({
                                                "header": tbl["header"],
                                                "rows": tbl["rows"],
                                                "slide_index": slide_idx,
                                            })
                                except Exception as ocr_err:
                                    st.warning(f"OCR failed on slide {slide_idx + 1}: {ocr_err}")

                        st.session_state["titles"][slide_idx] = title
                        st.session_state["bodies"][slide_idx] = body

                    st.session_state["slides_meta"] = slides_meta
                    st.session_state["tables"]      = table_data
                    st.session_state["input_file_uploaded"] = True
                    st.session_state["pptx_output_bytes"]   = None  # reset prior export

                    slides_with_ocr = sum(
                        1 for m in slides_meta
                        if not m.get("body_text", "").strip() and m.get("image_shapes")
                    )
                    st.success(
                        f"Analysis complete — **{len(slides_meta)} slides** extracted"
                        + (f", OCR applied to **{slides_with_ocr}** image slide(s)." if slides_with_ocr else ".")
                    )
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Review & Edit
# ══════════════════════════════════════════════════════════════════════════════
with tab_review:
    slides_meta = st.session_state.get("slides_meta", [])
    tables      = st.session_state.get("tables", [])

    if not slides_meta:
        st.markdown("""
        <div class="ss-empty">
            <div class="ss-empty-icon">🗂️</div>
            <div class="ss-empty-title">No slides analysed yet</div>
            <div class="ss-empty-desc">Upload a PPTX and run <strong>Analyse Presentation</strong> in the Upload tab.</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div class="ss-section">Slide Editor — {len(slides_meta)} slides</div>',
            unsafe_allow_html=True,
        )
        col_hint, col_save = st.columns([3, 1])
        with col_hint:
            st.caption("Expand each slide to edit its title and body text. Changes are saved automatically.")

        for meta in slides_meta:
            idx   = meta["slide_index"]
            title = st.session_state["titles"].get(idx, "")
            body  = st.session_state["bodies"].get(idx, "")
            tables_for_slide = [t for t in tables if t["slide_index"] == idx]

            # Status indicator in expander label
            has_body  = bool(body.strip())
            has_table = bool(tables_for_slide)
            status    = "📋" if has_table else ("✅" if has_body else "⚠️")

            with st.expander(f"{status}  Slide {idx + 1}  —  {title or '(no title)'}", expanded=False):
                new_title = st.text_input(
                    "Title", value=title, key=f"title_{idx}",
                    placeholder="Enter slide title…",
                )
                new_body = st.text_area(
                    "Body Text", value=body, height=180, key=f"body_{idx}",
                    placeholder="Body content will appear here after analysis…",
                )
                st.session_state["titles"][idx] = new_title
                st.session_state["bodies"][idx] = new_body

                if tables_for_slide:
                    st.markdown(
                        '<div style="font-size:0.72rem;font-weight:700;color:var(--muted);'
                        'text-transform:uppercase;letter-spacing:0.08em;margin:0.75rem 0 0.4rem;">'
                        'Extracted Table</div>',
                        unsafe_allow_html=True,
                    )
                    for tbl in tables_for_slide:
                        st.dataframe(tbl["rows"], use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Generate & Export
# ══════════════════════════════════════════════════════════════════════════════
with tab_export:
    slides_meta = st.session_state.get("slides_meta", [])

    if not slides_meta:
        st.markdown("""
        <div class="ss-empty">
            <div class="ss-empty-icon">📥</div>
            <div class="ss-empty-title">Nothing to export yet</div>
            <div class="ss-empty-desc">Complete the Upload &amp; Analyse step first.</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        col_gen, col_dl = st.columns([1, 1], gap="large")

        with col_gen:
            st.markdown('<div class="ss-section">Generate Output</div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:0.75rem;color:var(--muted);margin-bottom:1rem;line-height:1.7;">'
                'Assembles the final standardised PPTX from all reviewed slide content.<br>'
                'A brand template is required to preserve background styles.'
                '</div>',
                unsafe_allow_html=True,
            )

            if st.button("Generate Standardised PPTX", key="btn_generate", use_container_width=True):
                tables = st.session_state.get("tables", [])
                pages  = [
                    {
                        "title": st.session_state["titles"].get(m["slide_index"], f"Slide {m['slide_index']+1}"),
                        "body":  st.session_state["bodies"].get(m["slide_index"], ""),
                        "slide_index": m["slide_index"],
                    }
                    for m in slides_meta
                ]

                template_bytes = None
                try:
                    # template_ppt is defined in tab 1 but Streamlit re-runs keep the widget value
                    if "template_ppt_uploader" in st.session_state and st.session_state["template_ppt_uploader"] is not None:
                        template_bytes = st.session_state["template_ppt_uploader"].read()
                except Exception:
                    pass

                with st.spinner("Generating PPTX…"):
                    try:
                        output = template_filler.fill_template_with_pages(
                            template_bytes=template_bytes,
                            pages=pages,
                            tables=tables,
                            title_font=TITLE_FONT_NAME,
                            body_font=BODY_FONT_NAME,
                        )
                        st.session_state["pptx_output_bytes"] = output
                        st.success("PPTX generated successfully. Download it on the right.")
                    except Exception as e:
                        st.error(f"Generation failed: {e}")
                        st.code(tb.format_exc())

        with col_dl:
            st.markdown('<div class="ss-section">Download</div>', unsafe_allow_html=True)

            if st.session_state["pptx_output_bytes"]:
                size_kb = len(st.session_state["pptx_output_bytes"]) / 1024
                st.markdown(f"""
                <div class="ss-card" style="margin-bottom:1rem;">
                    <div class="ss-card-header">
                        <span class="ss-card-title">Output File</span>
                        <span class="ss-card-icon">📄</span>
                    </div>
                    <div style="font-size:0.82rem;font-weight:600;color:var(--text);">
                        standardised_output.pptx
                    </div>
                    <div style="font-size:0.7rem;color:var(--muted);margin-top:3px;">
                        {len(slides_meta)} slides · {size_kb:.1f} KB
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Determine a sensible file name
                try:
                    src_name = st.session_state["input_ppt_uploader"].name
                    out_name = f"standardised_{src_name}"
                except Exception:
                    out_name = "standardised_output.pptx"

                st.download_button(
                    label="Download Standardised PPTX",
                    data=st.session_state["pptx_output_bytes"],
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="download_pptx_button",
                    use_container_width=True,
                )
            else:
                st.markdown("""
                <div class="ss-empty" style="padding:2rem 1rem;">
                    <div class="ss-empty-icon">⏳</div>
                    <div class="ss-empty-title">Not generated yet</div>
                    <div class="ss-empty-desc">Click <strong>Generate Standardised PPTX</strong> on the left.</div>
                </div>
                """, unsafe_allow_html=True)
