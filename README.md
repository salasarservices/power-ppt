# PPT Normalizer — Phase B (MVP + OCR)

This repository contains a Streamlit app that extracts Title + Body from uploaded PPTX slides, optionally performs OCR on embedded images (Google Vision by default), provides an editable per-slide preview, paginates content, and fills a runtime template to generate a standardized PPTX.

Quick start
1. Create a virtual environment and install dependencies:
   pip install -r requirements.txt

2. Run the app:
   streamlit run app.py

Overview and behavior
- The app prefers to extract text from PPTX editable shapes (Title placeholder and body text shapes).
- If the extracted body is missing or very short and "Use OCR on images" is enabled, the app runs OCR on embedded image shapes detected in the slide (screenshots or pictures).
- Google Cloud Vision is the default OCR backend. Add your service account JSON to Streamlit Secrets using the key: GOOGLE_SERVICE_ACCOUNT_JSON
  - Alternatively, upload the service account JSON in the app sidebar for that session.
- Per-slide editable preview is shown after analysis so you can edit Title/Body before generation.
- Pagination:
  - Default: paragraph-based heuristic (chars per page).
  - Precise pagination: upload Poppins TTF at runtime to enable font-based measurements for pagination (optional).
- Continuation suffix: by default the app appends "(CONTD...)" to continuation pages' titles. You can edit this in the sidebar.

Important notes & limitations
- OCR currently operates on embedded images only (common case for screenshot text). The app does not rasterize full slides to images by default, because converting PPTX slides into images requires external tools (LibreOffice headless or PowerPoint automation) which may not be available in all hosting environments. If you want full-slide OCR, I can add that path; it requires installing additional system dependencies or invoking LibreOffice on the host.
- The app sets text font names to "Poppins" in generated slides. python-pptx cannot embed fonts; to see Poppins exactly, ensure the machine opening the final PPTX has Poppins installed.
- If you plan to use the app on Streamlit Cloud, add Google credentials to Streamlit Secrets rather than uploading in the UI for production.

Google Vision credentials (recommended)
- In Streamlit Cloud, open the Settings → Secrets and add a variable:
  GOOGLE_SERVICE_ACCOUNT_JSON = <<paste the full JSON contents>>
- The app will read this secret and initialize the Vision client.
- For local testing, you can:
  - Set the environment variable GOOGLE_APPLICATION_CREDENTIALS to the path of your JSON file, or
  - Upload the JSON in the app sidebar (session-only).

Next steps / improvements
- Add full-slide rasterization using LibreOffice headless for slide->image conversion to enable OCR on any slide, not just embedded images.
- Integrate TrOCR (Hugging Face) as a local, high-quality fallback (requires PyTorch).
- Improve layout detection using Detectron2 or Vertex AI custom detector for title/body region detection on complex slides.
- Add unit tests and more robust reading-order heuristics.

If you want me to proceed with any of these improvements (e.g., add full-slide rasterization), tell me and I will implement them next.
