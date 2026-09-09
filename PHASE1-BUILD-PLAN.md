# PowerPPT — Phase-1 Build Plan

_For sign-off before any code. Companion to [ROADMAP.md](ROADMAP.md). Drafted 09 Sep 2026._

**Review log (09 Sep 2026):** heading size set to **20pt**, body to **12pt**; the
**green rule under the heading is turned OFF** (`geometry.SHOW_GREEN_RULE = False`) —
a recorded deviation from the documented brand slide spec, reversible via that toggle.

## Phase-1 goal

Build the **brand engine** — the Python module that fills the authorised template —
as a clean, stack-agnostic package with golden tests. This is Package 1 (template-fit)
+ Package 2 (integrity guards) from the roadmap. No API, no UI, no deploy in this
phase; those follow once the engine is proven. The engine's I/O contract
(`SlidePlan`) is fixed here so the API, UI and v2 AI mode all plug into it unchanged.

---

## 1. Repo layout (monorepo — mirrors the Nexus `-api` / `-web` split, trimmed)

```
power-ppt/                         # github.com/salasarservices/power-ppt
  power-ppt-api/                   # FastAPI service (Phase 2 wraps this)
    app/
      main.py                      # FastAPI app + CORS (Phase 2)
      core/config.py               # pydantic-settings; Secret Manager in prod
      schemas/                     # Pydantic contracts (SlidePlan, Page, Table)
      brand_engine/                # <-- PHASE 1 DELIVERABLE
        __init__.py
        geometry.py                # measured EMU constants (this doc, section 3)
        template.py                # load + clone the template slide
        heading.py                 # brand heading: ALL CAPS, blue/green split, green rule
        content.py                 # body text + tables into the content zone
        builder.py                 # build_deck(slide_plan) -> bytes
        integrity.py               # verify_output(): the Package-2 guards
        errors.py                  # BrandEngineError
      extract/pptx_reader.py       # lifted from current repo (input extraction)
      ocr/                         # lifted: ocr_backend.py + preprocessor.py
      paginator.py                 # lifted + fixed (getsize -> getbbox)
    templates/2026/
      Salasar_Corporate_Blank.pptx # bundled, locked template (D2)
      VERSION                       # "2026"
    tests/
      fixtures/                    # sample input .pptx + a known SlidePlan
      test_brand_engine.py         # golden tests
      test_integrity.py
    pyproject.toml
    Dockerfile                     # Phase 2
    .env.example
  power-ppt-web/                   # Vite React (Phase 3)
  legacy-streamlit/                # current app.py etc. parked here, retired after cutover
  ROADMAP.md
  PHASE1-BUILD-PLAN.md
```

Existing modules are **reused, not rewritten**: `pptx_reader`, `ocr_backend`,
`preprocessor`, `paginator` move in as-is (paginator gets the Pillow fix).
`template_filler.py` is **replaced** by `brand_engine/` (it assumes a template
structure this year's file doesn't have).

---

## 2. Data contract (`SlidePlan`) — the seam every mode plugs into

Pydantic models in `app/schemas`. Reformat mode, and v2 AI Generate mode, both
produce a `SlidePlan`; the engine only ever sees this.

```python
class Table(BaseModel):
    header: list[str]
    rows:   list[list[str]]          # includes header as rows[0]

class Page(BaseModel):
    title:  str                      # may be "" -> engine renders no heading
    body:   str = ""                 # paragraphs separated by "\n\n"
    tables: list[Table] = []

class SlidePlan(BaseModel):
    pages: list[Page]
```

`build_deck(plan: SlidePlan, template_path: str) -> bytes` is the public entry point.

---

## 3. Geometry constants (`geometry.py`) — measured from the 2026 template

All values in inches; convert with `pptx.util.Inches`. Slide is 13.33 × 7.50.

```
MARGIN_L        = 0.75
MARGIN_R        = 0.75            # right edge = 12.58
LOGO_LEFT_EDGE  = 10.45          # heading must stay left of this

# Heading (top-left, clears the logo)
HEAD_LEFT   = 0.75
HEAD_TOP    = 0.55
HEAD_WIDTH  = 9.50               # 0.75 -> 10.25, clears logo
HEAD_HEIGHT = 0.90
RULE_GAP    = 0.08               # gap below heading text to the green rule
RULE_HEIGHT = 0.055
RULE_MIN_W  = 0.80               # fallback rule width if first-word measure fails

# Content zone (body / tables) — below heading, above footer bar (top 7.03)
BODY_LEFT   = 0.75
BODY_TOP    = 1.55
BODY_WIDTH  = 11.83              # 0.75 -> 12.58
BODY_BOTTOM = 6.83               # footer top 7.03 - 0.20 margin
BODY_HEIGHT = BODY_BOTTOM - BODY_TOP   # 5.28
```

Brand colours (hex): Blue `1A3A8F`, Green `7AC143`, Mid-Blue `0070C0`,
Slate `4D4D4D`, White `FFFFFF`.

---

## 4. Module specs

### 4.1 `template.py`
- `load_template(path) -> Presentation`.
- `clone_content_slide(prs) -> slide`: deep-copy the template's slide-0 shape tree
  (the group = background image + logo), copy relationships, remap rIds, re-number
  every `cNvPr` id to avoid duplicates. (Proven logic salvaged from
  `template_filler._copy_slide` + `_fix_shape_ids`.)
- The template has **no title placeholder and no "TITLE GOES HERE" marker** — so the
  old marker path is dropped entirely; the engine *adds* heading/body shapes onto the
  cloned canvas.

### 4.2 `heading.py` — full brand spec (D3)
`render_heading(slide, title)`:
1. If `title` blank → render nothing.
2. `text = title.strip().upper()` (ALL CAPS).
3. **Colour split:** split on the first hyphen (`-`, with or without surrounding
   spaces). `primary` (before) → Blue `1A3A8F`; `qualifier` (after) → Green `7AC143`;
   the hyphen stays with the primary run. No hyphen → whole title Blue.
4. Font: **Poppins Semi-Bold**, left-aligned, one line (auto-shrink if it would exceed
   `HEAD_WIDTH`). Textbox at `HEAD_LEFT/TOP/WIDTH/HEIGHT`.
5. **Green rule under the first word:** a filled rectangle, Green `7AC143`, height
   `RULE_HEIGHT`, at `HEAD_LEFT`, `HEAD_TOP + text_height + RULE_GAP`. Width = measured
   pixel width of the first word via the bundled Poppins TTF (Pillow `getbbox`), scaled
   to inches; fallback `RULE_MIN_W`.

### 4.3 `content.py`
`render_body(slide, body)`:
- Textbox at `BODY_LEFT/TOP/WIDTH/HEIGHT`, `word_wrap=True`.
- Poppins Regular, size 14pt, colour Slate `4D4D4D`, **line-height 1.6**, left-aligned,
  paragraphs preserved on `\n\n`. Content preserved verbatim (sentence case as-is — no
  forced casing on body).
- Overflow: if the body exceeds `BODY_HEIGHT`, `paginator` splits it across extra
  pages (continuation suffix on the title). Wired in this phase; `getsize`→`getbbox`
  fix applied.

`render_table(slide, table)`:
- `add_table` at content-zone top, row height ~0.34".
- Header row: fill Blue `1A3A8F`, text White `FFFFFF`, Poppins Semi-Bold 11pt.
- Body rows: Poppins Regular 10pt, Slate `4D4D4D`.

### 4.4 `builder.py`
`build_deck(plan, template_path) -> bytes`:
1. Load template; record its original slide count.
2. For each `Page` (after pagination): `clone_content_slide`, `render_heading`,
   `render_body` / `render_table`, strip any stray slide-number field.
3. Remove the original template slides.
4. Run `integrity.verify_output` on the saved bytes; raise on failure.
5. Return bytes.

### 4.5 `integrity.py` — Package-2 guards (a corrupt/incomplete file never returns)
`verify_output(pptx_bytes, plan) -> None` (raises `BrandEngineError` with detail):
- Re-opens with python-pptx (basic XML validity).
- Slide count == number of rendered pages.
- No duplicate `cNvPr` ids within any slide.
- **Brand elements present:** every slide contains the template's background + logo
  images (verify by image blob hash against the template's two pictures).
- **Content integrity:** each page's non-empty title text is found in its slide; a
  sampled body line from each page is present.

---

## 5. Test plan (golden tests — stops the regression cycle)

`tests/` with a small committed fixture deck:
1. `build_deck` on a known `SlidePlan` → output opens in python-pptx.
2. Slide count matches; no duplicate shape ids.
3. Background + logo images present on every slide (hash match).
4. Heading: primary run is Blue, qualifier run is Green, green rule shape exists.
5. Body stays within the content zone (top ≥ 1.55, bottom ≤ 6.83).
6. `verify_output` raises on a deliberately corrupted deck (negative test).
7. Pagination: an over-long body produces >1 page with the continuation suffix.

Run in CI (Phase 2). `pyproject.toml` pins dependencies (no unpinned ranges).

---

## 6. Explicitly NOT in Phase 1

API endpoints, FastAPI app, Vite/shadcn UI, auth/IAP, Cloud Run deploy, CI/CD,
GCS, database, v2 AI mode. Those are Phases 2–4 in the roadmap. Phase 1 ends with a
tested engine callable as `build_deck(plan, template_path)`.

---

## 7. Acceptance criteria

- `build_deck` turns a `SlidePlan` into a valid, on-brand PPTX against the 2026
  template: correct heading treatment, body in the content zone, footer + logo intact.
- All golden tests pass; `verify_output` blocks any malformed output.
- No hardcoded "TITLE GOES HERE" assumption remains.
- Dependencies pinned; engine has zero Streamlit/FastAPI/React imports (stack-agnostic).
