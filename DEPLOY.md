# PowerPPT — deploy & run runbook

Single **Cloud Run** service serving both the SPA and the API. Manual deploy
(no CI/CD yet). Project `power-ppt-486306`, region `asia-south1` (Mumbai).

---

## 0. One-time prerequisites (per machine)

```bash
gcloud auth login digitalmarketing@salasarservices.co.in
gcloud config set project power-ppt-486306
```

The deploy account needs `roles/owner` or, more narrowly, Cloud Run Admin +
Cloud Build Editor + Service Account User + Service Usage Admin.
(`digitalmarketing@` currently holds **Owner** on this project — verified 09 Sep 2026.)

Enable the APIs the deploy needs (idempotent — safe to re-run):

```bash
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com \
  secretmanager.googleapis.com \
  iap.googleapis.com
```

---

## 1. Deploy

Cloud Build builds the image from the repo-root `Dockerfile` (multi-stage:
Vite SPA -> FastAPI runtime). No local Docker needed. Run from the repo root:

```bash
gcloud run deploy power-ppt \
  --source . \
  --region asia-south1 \
  --no-allow-unauthenticated \
  --set-env-vars POWERPPT_WEB_DIST_PATH=/app/web_static
```

`--no-allow-unauthenticated` keeps the service private (no public anonymous
access) — required before putting IAP in front. First deploy takes a few minutes
(image build + push). It prints the service URL on success.

---

## 2. Lock it down with IAP

Cloud Run supports IAP directly (no external load balancer needed for an
internal tool).

1. Configure the OAuth consent screen once (Console -> APIs & Services -> OAuth
   consent screen), Internal user type, org `salasarservices.co.in`.
2. Enable IAP on the service:

   ```bash
   gcloud beta run services update power-ppt --region asia-south1 --iap
   ```

3. Grant the people who may use it:

   ```bash
   gcloud run services add-iam-policy-binding power-ppt \
     --region asia-south1 \
     --member="user:SOMEONE@salasarservices.co.in" \
     --role="roles/iap.httpsResourceAccessor"
   ```

   (Use `group:` for a Google Group instead of adding users one by one.)

After IAP is on, the service is reachable only to granted org members through
the Google sign-in gate — matching the estate's internal-tool pattern.

---

## 3. Secrets (only if server-side OCR keys are used)

The brand engine needs no secrets. OCR via Google Vision should use the Cloud
Run **runtime service account** (ADC) — grant it Vision access rather than
pasting a key. If a key file is genuinely required:

```bash
echo -n "$SERVICE_ACCOUNT_JSON" | gcloud secrets create powerppt-vision-sa --data-file=-
gcloud run services update power-ppt --region asia-south1 \
  --set-secrets POWERPPT_GOOGLE_SERVICE_ACCOUNT_JSON=powerppt-vision-sa:latest
```

---

### Document AI (A3 — image-of-a-table → native editable table)

Auth is the Cloud Run runtime service account (ADC) — no key file.

1. Enable the API:
   ```bash
   gcloud services enable documentai.googleapis.com
   ```
2. Create a **Form Parser** processor (Console → Document AI → Create Processor →
   Form Parser), region `us` or `eu`. Copy the **Processor ID**.
3. Grant the runtime service account access:
   ```bash
   gcloud projects add-iam-policy-binding power-ppt-486306 \
     --member="serviceAccount:400070465780-compute@developer.gserviceaccount.com" \
     --role="roles/documentai.apiUser"
   ```
4. Point the service at the processor:
   ```bash
   gcloud run services update power-ppt --region asia-south1 --set-env-vars \
     POWERPPT_DOCAI_PROJECT=power-ppt-486306,POWERPPT_DOCAI_LOCATION=us,POWERPPT_DOCAI_PROCESSOR_ID=<PROCESSOR_ID>
   ```

When set, an uploaded image that Document AI detects as a table is converted to a
native editable table and the image is dropped; other images pass through unchanged.
Unset → images are always kept as-is (no table OCR). Billed per page.

### Tier-2 AI extraction — Vertex AI Gemini (picture / freeform decks)

For decks whose text/tables are baked into pictures or vector art (native extraction
can't read them), the API renders the deck to PDF (LibreOffice, already in the image)
and asks **Vertex AI Gemini** to transcribe each slide, merging only where native
came up short. **Compliance: Vertex only (enterprise, no-train); the model is
instructed to transcribe, never invent. Client decks are permitted on Vertex; free-tier
Gemini/AI Studio is not.**

1. Enable Vertex AI:
   ```bash
   gcloud services enable aiplatform.googleapis.com
   ```
2. Grant the runtime service account:
   ```bash
   gcloud projects add-iam-policy-binding power-ppt-486306 \
     --member="serviceAccount:400070465780-compute@developer.gserviceaccount.com" \
     --role="roles/aiplatform.user"
   ```
3. Point the service at Vertex + turn it on:
   ```bash
   gcloud run services update power-ppt --region asia-south1 --set-env-vars \
     POWERPPT_VERTEX_PROJECT=power-ppt-486306,POWERPPT_VERTEX_LOCATION=us-central1,POWERPPT_VERTEX_MODEL=gemini-2.5-flash,POWERPPT_AI_EXTRACT_DEFAULT=true
   ```

Per-request override: `POST /analyze?use_ai=true` (or `false`). Billed per deck by
token/page; renders + one Gemini call per upload. Unset/unconfigured → native only.

## 4. Run locally (dev)

Two processes: FastAPI on :8077, Vite on :5173 (Vite proxies the API routes).

```bash
# terminal 1 — API
cd power-ppt-api && python -m uvicorn app.main:app --port 8077

# terminal 2 — web
cd power-ppt-web && npm install && npm run dev
```

Open http://localhost:5173. Tests: `cd power-ppt-api && python -m pytest -q`.

### Verify the single-service build locally (optional, no Docker)

```bash
cd power-ppt-web && npm run build            # -> power-ppt-web/dist
cd ../power-ppt-api
POWERPPT_WEB_DIST_PATH="$(pwd)/../power-ppt-web/dist" \
  python -m uvicorn app.main:app --port 8078
# http://localhost:8078/ serves the SPA; /health serves the API.
```

---

## 5. Rollback

Cloud Run keeps every revision. Roll back by shifting traffic:

```bash
gcloud run services update-traffic power-ppt --region asia-south1 --to-revisions=PREVIOUS=100
```

List revisions: `gcloud run revisions list --service power-ppt --region asia-south1`.
