# MatdataAI

**Convert Hindi voter-list PDFs into structured, review-ready Excel data — automatically.**

Built for SUMMER SCHOOL'26 (AI First Hackathon) by Team **The Solvers**.

---

## 🧩 Problem

Booth-Level Officers (BLOs), candidates, and party workers routinely need voter-roll data
(names, EPIC numbers, house numbers, photos) from official Hindi PDF voter lists in a
structured, searchable format. Today this is done by **manual data entry**, which is slow,
error-prone, and doesn't scale to lakhs of records. MatdataAI automates this end-to-end.

---

## 🧠 AI Architecture

The pipeline is a 3-phase system that takes a raw voter-list PDF and produces a
structured Excel file with embedded photos and a separate sheet for low-confidence entries
that need human review.

```
                ┌─────────────────────┐
   PDF file --> │  Phase 1: Crop       │  OpenCV contour detection to locate each
                │  (OpenCV +           │  grid-cell/photo box on the page, then
                │   insightface)       │  insightface confidence-check to verify
                └─────────┬────────────┘  a valid face photo was actually cropped.
                          │
                          v
                ┌─────────────────────┐
                │  Phase 2: Extract    │  Tesseract Hindi OCR reads the text fields
                │  (Tesseract Hindi    │  next to each photo, a custom parser maps
                │   OCR + parser)      │  raw OCR text into structured fields
                └─────────┬────────────┘  (name, relation's name, house no., age,
                          │               gender, EPIC no.), flags low-confidence
                          │               reads instead of silently guessing.
                          v
                ┌─────────────────────┐
                │  Phase 3: Excel Gen  │  openpyxl writes one row per voter with the
                │  (openpyxl)          │  embedded photo, plus a separate "Review"
                └─────────┬────────────┘  sheet for entries flagged in Phase 2.
                          │
                          v
                  Structured .xlsx output
```

**Why this replaced the earlier Gemini-API version:** the original prototype called
Gemini 2.5 Flash for OCR/parsing. This version uses local, open-source models
(OpenCV + insightface + Tesseract) instead — no per-request API cost, no external
rate limits, and it works fully offline once dependencies are installed. This matters
at scale: the target dataset is **362 PDFs / 11,452 pages**, which would be
cost-prohibitive and rate-limited on a pay-per-call API.

**Low-confidence handling:** rather than silently guessing on garbled OCR output, flagged
entries are routed to a dedicated review sheet — accuracy claims are validated against a
stratified random audit sample pulled across *all* processed files (not just the initial
test batch), not just eyeballed on a handful of PDFs.

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Photo cropping / face verification | OpenCV (contour detection), insightface |
| OCR / text extraction | Tesseract OCR (Hindi language pack) |
| Field parsing | Custom Python parser |
| Excel generation | openpyxl, pandas |
| Backend API | FastAPI, Uvicorn |
| Batch processing / checkpointing | Custom Python batch runner (JSON manifest, resumable) |
| Frontend | HTML5, CSS3, vanilla JavaScript |
| Hosting — backend | Render (Docker web service) |
| Hosting — frontend | GitHub Pages |

---

## 💻 Local Setup Instructions

### Backend (processing API)

```bash
git clone <this-repo-url>
cd <repo-folder>/backend

# Install system dependencies (Tesseract with Hindi pack)
# On Ubuntu/Debian:
sudo apt-get install tesseract-ocr tesseract-ocr-hin libgl1 libglib2.0-0

# Install Python dependencies
pip install -r requirements.txt

# Run the API locally
uvicorn backend_api:app --reload --port 8000
```

The API will be live at `http://localhost:8000`. Interactive API docs (Swagger) are at
`http://localhost:8000/docs`.

**Or, using Docker (recommended — matches the deployed environment exactly):**

```bash
docker build -t matdataai-backend .
docker run -p 8000:8000 matdataai-backend
```

### Frontend (website)

The website is a self-contained static HTML file — no build step required.

```bash
cd frontend
# Just open index.html in a browser, or serve it locally:
python -m http.server 5500
```

Then open `http://localhost:5500` in your browser. Update the `API_BASE` constant near
the top of the `<script>` block to point at your local backend
(`http://localhost:8000`) or the deployed Render URL.

### Running the full batch pipeline (bulk dataset processing)

For processing the full 362-PDF dataset (not needed to test the demo, only relevant for
the bulk-processing use case):

```bash
python voter_pipeline_batch_runner.py
```

See inline comments in that file for the checkpoint/resume behavior across sessions.

---

## 🌐 Live Demo

- **Website:** `<your-github-pages-url>`
- **Backend API:** `<your-render-url>` (interactive docs at `/docs`)
- **Demo video:** `<your-video-link>`

> Note: the backend runs on a free-tier instance and may take 30–60 seconds to respond
> on the very first request after a period of inactivity (cold start). Subsequent
> requests are fast.

---

## 📈 Feasibility & Scalability

The full target dataset is **362 PDF files, 11,452 pages** (~19,500 entries per 25-PDF
sample, extrapolating to ~280,000+ entries dataset-wide), sourced from Google Drive.
Processing this reliably — especially on free-tier compute with frequent session
disconnects (e.g. Colab) — required building:

- **Batching** — PDFs are processed in tunable-size chunks per session rather than all
  at once, sized using measured per-page processing time so each run finishes within a
  safe window.
- **Checkpoint/resume** — a JSON manifest tracks `done` / `failed` / `in_progress` PDFs.
  If a session disconnects mid-run, re-running the script picks up exactly where it left
  off — no PDF is reprocessed unnecessarily.
- **Graceful error handling** — corrupted PDFs, transient Drive-read glitches, and
  unexpected page layouts are caught and logged per-file/per-page; one bad file never
  crashes the full run.
- **Validation at scale** — a stratified random audit sample is drawn across *all*
  processed files (not just the original test batch) to measure real dataset-wide
  accuracy rather than relying on the small initial sample.

This design means the same pipeline that works on a 25-PDF test scales to the full
dataset without manual babysitting of every run.

---

## 👥 Team — The Solvers

Built by [your name(s) here].
