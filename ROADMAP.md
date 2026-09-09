# PowerPPT — Roadmap & Scope

_Owner: DevOps lead · Status: stabilisation phase · Last updated: 09 September 2026_

Internal planning document. Not client-facing.

---

## ▶ Resume here (state as of 09 Sep 2026)

**Fresh-session start:** read this section, then the Target architecture + Sequencing
sections below, and `PHASE1-BUILD-PLAN.md`. Repo: `github.com/salasarservices/power-ppt`.
GCloud project: `power-ppt-486306`.

**Done & pushed to `main`** (commit `6724983`):
- ✅ Phase 1 — brand engine (`power-ppt-api/app/brand_engine/`): fits the 2026
  template, heading (blue/green split, green rule OFF, heading 20pt / body 12pt),
  body + tables, integrity guards. Retired the old Streamlit app to `legacy-streamlit/`.
- ✅ Phase 2 — FastAPI service (`power-ppt-api/app/`): `/health`, `/generate`,
  `/analyze`; OCR lazy-loaded; Dockerfile + env settings. **17 tests pass.**

- ✅ Phase 3 — Vite/shadcn web UI (`power-ppt-web/`): upload → review → export flow
  against the live API. React 18 + TS + Vite + Tailwind + hand-rolled shadcn-style
  primitives, TanStack Query + axios. Brand shell (Poppins, brand colours, ImageKit
  logo, gradient footer bar). Verified end-to-end in-browser (analyze → edit plan →
  generate → download) on 09 Sep 2026. `npm run build` passes.

**Phase 4 — in progress (artifacts ready, not yet deployed):**
- ✅ Topology decided: **single Cloud Run service** — FastAPI serves the built SPA
  (`web_static/`) at `/` and the API at root; conditional static mount keeps the 17
  tests green. Verified locally (SPA at `/`, `/health` API wins). CI/CD deferred
  (manual deploy first).
- ✅ Deploy artifacts: repo-root multi-stage `Dockerfile` (Vite build → FastAPI),
  `.dockerignore`, `power-ppt-api/.env.example`, and `DEPLOY.md` runbook
  (enable-APIs → `gcloud run deploy --source .` region `asia-south1` → IAP → secrets).
- ✅ Access confirmed: `digitalmarketing@` holds **roles/owner** on `power-ppt-486306`;
  project ACTIVE; deploy APIs (run/cloudbuild/artifactregistry/secretmanager/iap) NOT
  yet enabled (23 BigQuery-era services on).
- ⬜ **Remaining (user runs — needs interactive `gcloud auth login`):** enable APIs,
  `gcloud run deploy`, enable IAP + grant users. No CI/CD workflow yet (by design).

**Open decision (user to choose):** deploy the API now for a live URL, OR build the UI
first then deploy web+api together, OR just add a CI test workflow. User dismissed this
choice on 09 Sep — awaiting direction next session.

**Run locally:** `cd power-ppt-api && python -m pytest -q` (tests) ·
`python -m uvicorn app.main:app --port 8077` then hit `/health`, `/docs`.
Note: dev machine is Windows + Python 3.14; keep Pillow ≥12.2 (global `pdfplumber` needs it).

---

## Purpose

PowerPPT reformats arbitrary PowerPoint decks into the authorised Salasar brand
template, keeping the source content intact. It extracts titles, body text,
tables and image content (via OCR when a slide is image-only), then re-pours that
content into the locked brand template so every output carries the correct logo,
footer bar, colours and Poppins typography.

Fundamental principle: **the app formats content, it does not author facts.** Any
future generation feature (see v2) feeds the same brand engine and never bypasses
the compliance gate.

---

## Decisions on record

GCloud project for deployment: **`power-ppt-486306`** (exists) under org
`salasarservices.co.in`.

| # | Decision | Rationale |
|---|----------|-----------|
| D1 | Stabilise the current stack first; no rewrite | Core is broken against the new template today; the value lives in the Python `python-pptx` engine, which any stack keeps. |
| D2 | Bundle the authorised template in the repo as the locked default | Template is authorised, fixed for the year, changes annually. Removes "wrong template" errors; matches brand governance. |
| D3 | Heading rendering: full brand spec | ALL CAPS, blue primary term + green qualifier after hyphen, green rule under the first word. |
| D4 | **Re-platform off Streamlit** onto the standard Google Cloud web-app stack, to match Salasar's other GCloud apps | Org standardisation: shared frontend conventions, team skills, one ops/IAM/billing model. **Invariant:** the PPTX engine stays Python (`python-pptx` has no JS equivalent) — so the frontend matches the house stack while a Python **FastAPI service on Cloud Run** does the deck engine, OCR and Gemini. House stack **confirmed** = React+Vite+shadcn (like `nexus-web`) + FastAPI (like `nexus-api`). See Target architecture. |
| D5 | AI "Generate" mode is v2, guardrailed | Useful adjacency, but hallucination and free-tier data-training conflict with IRDAI and client-data rules. See v2 section. |
| D6 | **Minimal footprint** — take from Nexus only the UI + tech-stack *choices*, not its weight | PowerPPT is a stateless convert tool, not a CRM. No DB / SQLAlchemy / Alembic / RBAC / audit tables / rate limiter. Files stream in-memory; auth via IAP (no app code); only the shadcn components actually used. Reuse Nexus's `tailwind.config` + Poppins tokens for consistent looks. |

### Authorised template — structure notes (measured 09 Sep 2026)

Source: `Salasar Corporate PPT_Blank Template_NEW.pptx` (fixed for this year).

- 13.33 × 7.5 in widescreen; 2 byte-identical slides.
- Each slide = one **group** of two images: a full-bleed white background with the
  green→blue **footer bar** baked in (rounded top corners, inset), and the **logo**
  top-right (L≈10.45, T≈-0.28, W≈2.63, H≈1.86 in).
- **No title placeholder, no `"TITLE GOES HERE"` marker, no text of any kind.**
  The current `template_filler.py` assumes both exist — hence titles are silently
  dropped against this template. This is the headline defect Package 1 fixes.

---

## Target architecture — mirror the confirmed house stack (`nexus-web` + `nexus-api`)

Verified 09 Sep 2026 by inspecting the Nexus repo. PowerPPT mirrors it
component-for-component so it looks and deploys like the rest of the estate. Repo of
record: `github.com/salasarservices/power-ppt`.

Minimal footprint (D6): take Nexus's **UI + stack choices only**, not its CRM weight.

| Layer | Choice | Role |
|-------|--------|------|
| Frontend | **React 18 + TypeScript + Vite** SPA · **Tailwind** + **shadcn/ui** (only the components used: button, card, input, textarea, tabs, table, scroll-area, sonner) · **TanStack Query** + **axios** · **react-router-dom** · **@fontsource/poppins** | Upload → review → export UI. |
| Backend | **Python FastAPI** on **Cloud Run** — thin | Brand engine (`python-pptx`), OCR, Gemini (v2). 3 endpoints: `/analyze`, `/generate`, `/health`. No JS equivalent for the deck engine. |
| Shared looks | Reuse **`nexus-web`'s `tailwind.config` + shadcn setup + Poppins tokens** | The standardised-looks mechanism — same tokens/components as Nexus. Extract a shared `@salasar/ui` package only once a second app needs it. |
| Auth | **IAP** on the Cloud Run service — no app-level auth code | Internal tool; infra-level protection is enough. |
| Secrets | **Secret Manager** for OCR/Gemini keys | The one piece of GCloud plumbing genuinely required. |
| State/Storage | **None for v1** — files stream in-memory (request → response) | Stateless convert. Add **GCS** signed URLs only if file size/timeouts force async. Audit = Cloud Logging, not a DB. |
| Ops | **Docker → Cloud Run**; structured logging; CORS locked to the frontend origin | Matches the estate without the extra tiers. |

**Deferred until actually needed** (not in v1): SQLAlchemy · Alembic · Postgres/Neon ·
RBAC · audit tables · slowapi rate limiting · GCS · react-hook-form · zod · Sentry.

Compliance is enforced **server-side** (banned-term filter, disclaimers, `[TO CONFIRM]`,
review gate) so the frontend cannot bypass it.

_Correction on record: an earlier draft recommended Next.js **and** a full Nexus-weight
backend; inspecting `nexus-web` fixed the framework to a **Vite React SPA with
shadcn/ui**, and D6 trimmed the backend to a thin stateless service._

Sequencing note: **go straight to the new build** (no Streamlit stopgap — team waits).
De-risk by shipping the Python brand engine first (independently testable), then the
thin FastAPI API, then the Vite/shadcn UI.

---

## Package 1 — Template-fit _(first, standalone — build as a stack-agnostic Python module)_

Make the generator compatible with the authorised image-canvas template.

- [ ] Bundle + load the locked template from `templates/2026/`; remove the
      "optional template" contradiction in the UI.
- [ ] Heading — full brand spec: top-left, ALL CAPS, primary term blue `#1A3A8F`,
      qualifier after hyphen green `#7AC143`, green rule under the first word,
      Poppins Semi-Bold. (Fallback when no hyphen: whole title blue; rule under
      first word regardless.)
- [ ] Body — brand spec: Poppins Regular, slate `#4D4D4D`, line-height 1.6,
      left-aligned, sentence case, kept clear of the footer bar.
- [ ] Retire the `"TITLE GOES HERE"` assumption; drive placement from measured
      template geometry (EMU units, via the `power-ppt` skill).
- [ ] **Design the seam:** brand engine accepts a generic slide-plan
      (`slides: [{title, body|bullets, tables}]`) so v2 generation drops in later
      without rework.
- [ ] Visual verification: generate a sample deck; confirm heading, body, footer
      and logo land correctly.

## Package 2 — Correctness & integrity guards (P0)

- [ ] Post-generation integrity check: re-open output; assert slide count, no
      duplicate shape IDs, valid XML. A corrupt file must never reach a user
      silently.
- [ ] Content-integrity guard: every input slide's text provably lands in output.

## Package 3 — Honesty, hygiene & deploy (P1/P2)

- [ ] Wire up pagination (`paginator.py`) or remove its dead UI controls
      (continuation suffix, chars-per-page, Poppins TTF are currently captured but
      unused).
- [ ] Fix `paginator.font.getsize` → `getbbox`/`getlength` (removed in Pillow 10+).
- [ ] Surface OCR failures instead of silently returning empty text.
- [ ] Pin dependencies (lockfile); add 3–4 golden tests (known input → assert
      output opens, slide count, brand elements present, no duplicate IDs).
- [ ] Short runbook: how the team runs it, where secrets live.
- [ ] **Re-platform to Google Cloud (supersedes "no frontend change"):** split into
      a **Python FastAPI backend** on **Cloud Run** (brand engine, OCR, Gemini) and
      a frontend on Salasar's standard GCloud web-app stack **[TO CONFIRM]**;
      secrets in **Secret Manager**, files in **GCS**, auth via **IAP / Identity
      Platform**. Retire Streamlit. The Package 1–2 engine is written stack-agnostic
      so it lifts into the FastAPI service unchanged.

---

## v2 — AI "Generate" mode _(future, guardrailed — not part of stabilisation)_

Let a user create a branded deck from bare information, in addition to reformatting.
The AI produces a **structured slide-plan only**; it feeds the same Package 1 brand
engine and the same review gate. The AI never gets the last word.

**Model:** Gemini (Google-native — aligns with Vision + Cloud Run).

**Non-negotiable guardrails (IRDAI + client-data rules):**

1. **Restructure, don't author facts.** The AI reorganises content the user
   supplies; it does not fetch coverage, exclusions, limits, premiums, dates or
   statistics from model knowledge. Missing facts render as `[TO CONFIRM]`.
2. **Data governance splits by content:**
   - Free-tier Gemini (AI Studio) — **only** for non-client, non-factual,
     structural drafting ("lay out these bullets I typed"). Free tiers may train on
     inputs, which conflicts with the client-data policy.
   - **Vertex AI on Salasar's GCP** (enterprise, no-train) — required for anything
     touching client or product facts.
   - Never send a client identifier to any LLM endpoint.
3. **Compliance filter on output:** block banned terms (guaranteed, assured
   returns, 100% claim settlement, cheapest, best in the business, risk-free,
   instant approval, no rejection, unlimited cover, we settle your claim, free);
   no unsourced statistics.
4. **Auto-append** entity name + IRDAI Licence No. 143, and "subject to policy
   terms, conditions and exclusions" on benefit statements.
5. **Mandatory human review** before any output reaches a client, insurer or
   regulator.

**Deferred infra, add only when a real need appears:** Postgres/Neon audit trail,
GCS async handoff, RBAC, Sentry — pull these in from the Nexus pattern if/when file
sizes, formal audit, or multi-role access actually require them.

---

## Sequencing

Decision: **go straight to the new GCloud build — no Streamlit stopgap.** The team
waits for the re-platform rather than investing in throwaway Streamlit UI.

1. ✅ **Brand engine** — Package 1 (template-fit) + Package 2 (integrity guards),
   stack-agnostic Python module. 12 golden tests. (`app/brand_engine/`)
2. ✅ **Thin FastAPI service** wrapping the engine (`/analyze`, `/generate`,
   `/health`); OCR lazy-loaded; Dockerfile + settings. 5 API tests. (`app/api/`)
3. ✅ **Vite/shadcn UI** (upload → review → export) against the API (`power-ppt-web/`).
   Brand tokens applied directly (no local `nexus-web` source found on this machine —
   only Nexus docs). Router/unused shadcn parts trimmed per D6. 5 primitives, 2 flow
   components, typed API client mirroring the FastAPI schemas.
4. **Deploy** to `power-ppt-486306` behind IAP; Docker → Cloud Run; CI/CD.  ← IN PROGRESS
   Single-service topology; artifacts + runbook (`DEPLOY.md`) ready. Live deploy +
   IAP pending (user runs the `gcloud` steps). CI/CD deferred.
5. Package 3 hygiene folded in (pagination fix ✅, dep pinning ✅, golden tests ✅,
   runbook).
6. **v2 — AI Generate mode** (Gemini, guardrailed), from the stable deployed base.

Stack lock: **resolved** — mirrors Nexus (`nexus-web` React+Vite+shadcn, `nexus-api`
FastAPI). Deploy target `power-ppt-486306`.
