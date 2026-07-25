"""
backend_api.py
================
FastAPI service that wraps the existing voter-roll pipeline (process_pdf
from voter_pipeline_batch_runner.py) so the MatdataAI website can call it
over HTTP. Deploy this SEPARATELY from the static website (see chat for
where to host it) — the website just talks to it over the internet.

Run locally to test:
    uvicorn backend_api:app --reload --port 8000

Endpoints:
    POST /api/process        -> upload a PDF, get back a job_id
    GET  /api/status/{job_id} -> {"status": "processing"|"done"|"failed", ...}
    GET  /api/download/{job_id} -> the generated .xlsx file
"""

import uuid
import shutil
import threading
from pathlib import Path
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

# your existing pipeline function
from voter_pipeline_batch_runner import process_pdf

UPLOAD_DIR = Path("uploads")
RESULT_DIR = Path("results")
UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="MatdataAI Processing API")

# CORS: allow requests from your GitHub Pages domain.
# Replace with your actual site URL(s) — never leave "*" in production
# once this is public, since it's an expensive endpoint (OCR compute).
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://<your-username>.github.io",   # <-- your GitHub Pages URL
        "http://localhost:5500",                # local testing (Live Server etc.)
    ],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

# in-memory job store. For a hackathon demo this is fine; for anything
# longer-lived, swap this dict for a small SQLite table or Redis.
jobs = {}


@app.get("/")
async def root():
    return {"status": "MatdataAI backend is running"}


def _run_job(job_id: str, pdf_path: Path):
    jobs[job_id]["status"] = "processing"
    try:
        # process_pdf() returns a dict:
        # {"main_path":..., "review_path":..., "n_entries":..., "n_flagged":...}
        result = process_pdf(str(pdf_path), output_dir=str(RESULT_DIR / job_id))

        excel_dst = RESULT_DIR / f"{job_id}.xlsx"
        shutil.copy(result["main_path"], excel_dst)

        review_dst = None
        if result.get("review_path"):
            review_dst = RESULT_DIR / f"{job_id}_review.xlsx"
            shutil.copy(result["review_path"], review_dst)

        jobs[job_id].update({
            "status": "done",
            "entries": result["n_entries"],
            "flagged": result["n_flagged"],
            "excel_path": str(excel_dst),
            "review_path": str(review_dst) if review_dst else None,
            "finished_at": datetime.utcnow().isoformat(),
        })
    except Exception as e:
        jobs[job_id].update({"status": "failed", "error": str(e)})


@app.post("/api/process")
async def process_pdf_endpoint(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Only PDF files are accepted")

    job_id = str(uuid.uuid4())
    saved_path = UPLOAD_DIR / f"{job_id}.pdf"
    with open(saved_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    jobs[job_id] = {"status": "queued", "created_at": datetime.utcnow().isoformat()}

    # run in background thread so the upload request returns immediately —
    # website then polls /api/status instead of holding the connection open
    thread = threading.Thread(target=_run_job, args=(job_id, saved_path), daemon=True)
    thread.start()

    return {"job_id": job_id, "status": "queued"}


@app.get("/api/status/{job_id}")
async def check_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Unknown job_id")
    return jobs[job_id]


@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    job = jobs.get(job_id)
    if not job or job["status"] != "done":
        raise HTTPException(404, "Result not ready or job not found")
    return FileResponse(
        job["excel_path"],
        filename="matdataai_output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/download/{job_id}/review")
async def download_review(job_id: str):
    job = jobs.get(job_id)
    if not job or job["status"] != "done" or not job.get("review_path"):
        raise HTTPException(404, "No review sheet for this job")
    return FileResponse(
        job["review_path"],
        filename="matdataai_review.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
