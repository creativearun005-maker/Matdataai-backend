# MatdataAI — मतदाताAI

**Convert scanned Hindi voter-list PDFs into structured, filterable Excel data — ward-wise, house-wise, near-zero cost.**

Built for Bharat's electoral data digitization at scale, replacing manual data-entry and closed-API dependency (Gemini) with an open, cost-controlled pipeline.

---

## The Problem

Election Commission of India voter lists are distributed as scanned/rasterized PDFs — 15 voter entries per page, printed Devanagari text, passport photos, and (in later batches) handwritten annotations. There is no official structured/machine-readable export. Candidates, BLOs (Booth Level Officers), and researchers currently do this extraction by hand.

## What MatdataAI Does

Given a folder of voter-list PDFs, MatdataAI outputs a clean Excel file per constituency part, with:

- Cropped, aligned voter photograph per row
- Extracted fields: EPIC number, name, father's/husband's name, house number, age, gender
- A **separate review sheet** for any field the system isn't confident about — with the source crop image attached, so a human reviewer can confirm/correct in seconds rather than re-reading the whole document
- (In progress) Handwritten mobile-number extraction from later-batch documents

## Architecture

The core design principle: **don't guess, flag.** Every extracted field either meets a confidence bar or is routed to a human-reviewable queue with its source image attached — never silently wrong, never silently dropped.

```
PDF → page-classification → photo crop → text-region crop → OCR → field-parse → validate
                ↓                ↓              ↓                      ↓
         (column-count      (contour +    (Tesseract Hindi      (regex + label-
          detection)         face-conf.)    OCR)                 similarity match)
                                                                        ↓
                                                          clean → Excel  |  flagged → Review sheet (+ crop image)
```

**Stage 1 — Page classification.** Each PDF mixes cover pages, polling-station-photo pages, the voter grid, and summary pages. Rather than assuming a fixed page layout, we detect the 3-column voter-grid structure directly (contour/column analysis) — this generalizes across documents with a different number of intro/summary pages, and correctly handles partial last-pages.

**Stage 2 — Photo cropping.** Voter photo-boxes are located via OpenCV contour detection on the printed cell-border (not the face itself) — this is deliberately more robust than face-detection-first, because scan quality varies enormously across this dataset (halftone/photocopy degradation) and border-ink survives that degradation far better than facial detail does. A face-detection pass (InsightFace) runs as a secondary confidence signal, not the primary detector — validated against 19,485 real crops with a 0.046% failure rate.

**Stage 3 — Text extraction.** Tesseract (Hindi trained data) reads the text region derived relative to each photo-box position (not fixed absolute coordinates — this keeps extraction robust to minor page-to-page layout drift). Fields are parsed by a position-and-similarity-based labeller (first line is always the name; subsequent lines matched against known field labels via string similarity, tolerant of OCR noise in the label itself).

**Stage 4 — Validation & review routing.** Numeric fields go through a character-confusion cleanup pass (OCR commonly confuses `1`/`]`/`|`, `0`/`O`, etc.). Any entry with an unresolved parsing issue is flagged, and its source crop is saved for the review sheet — this is generated automatically, not a manual step.

**Stage 5 — Excel generation.** `openpyxl`-based generation with embedded (not just linked) images, matching the required 15-column schema, plus the auto-generated review workbook.

## Tech Stack

| Layer | Tool | Why |
|---|---|---|
| PDF → image | PyMuPDF (`fitz`) | Fast, reliable rasterization, no external binary dependency |
| Photo detection | OpenCV (contour/edge detection) | Robust to scan-quality variance; doesn't depend on facial clarity |
| Face confidence | InsightFace | Secondary confidence signal only, not primary detector |
| OCR | Tesseract (`hin` trained data) | Best free option tested for printed Devanagari at this scan quality |
| Numeric field cross-check | Mistral AI API | Used selectively for high-stakes numeric fields (EPIC number) where Tesseract's digit-confusion rate was unacceptable — see *Known Limitations* |
| Excel output | `openpyxl` | Native embedded-image support, required for the photo column |
| Backend | Python API (deployed on Render) | See *Deployment Notes* below for the resource-tradeoff made here |

## Known Limitations & Deliberate Trade-offs (read this before judging accuracy)

We are stating these directly rather than hiding them:

1. **Digit-recognition accuracy on printed numerals is currently the weakest link**, specifically for thin-stroke digits (`1`) in this document's font — Tesseract intermittently drops them entirely rather than misreading them, which is a font-specific rendering issue, not a generic OCR-tuning problem. We've mitigated this for the EPIC-number field using a secondary API-based cross-check (Mistral), because EPIC number accuracy is a higher-priority correctness requirement than any other field. **Round 3 plan:** fine-tune an OCR model on this exact document font using our own scanned page corpus, removing the API dependency entirely.
2. **Handwritten-field extraction (mobile numbers) is at proof-of-concept stage**, not production-accuracy — see the Phase 5 section of the demo video for what's implemented vs. planned.
3. **The hosted demo (Render free tier) runs a reduced-resource version of the pipeline** — the face-confidence-check stage is disabled there due to free-tier RAM limits. The full pipeline (including face-confidence validation) is demonstrated separately running on Google Colab in the demo video, showing the accuracy numbers this architecture actually achieves with adequate resources. This is a hosting-cost decision, not an architecture limitation — the same code runs both places.
4. **~25-30% of entries currently require human review** (flagged, not wrong) on our validated 19,485-entry test sample. We consider this an honest, working safety-net rather than a shortcoming — the alternative (no flagging) would silently ship incorrect voter data, which is unacceptable for this use case.

## Local Setup

```bash
git clone <repo-url>
cd matdataai-backend
pip install -r requirements.txt
apt-get install tesseract-ocr tesseract-ocr-hin poppler-utils

# Run against a folder of voter-list PDFs
python voter_pipeline_batch_runner.py --input <pdf-folder> --village-number <name>
```

Output: `output/<pdf-name>.xlsx` (main) and `output/<pdf-name>_review.xlsx` (flagged entries + crop images).

## Live Demo

**[live demo link here]**

Note: hosted on Render's free tier — see *Known Limitations* above regarding the resource-reduced configuration running there.

## Demo Video

**[video link here]**

## Roadmap

- **Round 3:** Fine-tune OCR model on this document's specific font/print-style (own-corpus training, removing external API dependency for numeric fields), complete handwritten-field pipeline, EPIC-format validation.
