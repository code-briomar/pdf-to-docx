# PDF → Word Converter

A fully local, offline tool that converts any PDF — text-based, scanned, or
mixed — into an accurately formatted `.docx` file. Detects whether OCR is
needed on a per-page basis, extracts text/tables/structure, and rebuilds it
as a Word document. No API keys, no cloud dependency, no per-page cost.

## How it works

```
PDF in
  │
  ▼
Per-page check: does this page have a text layer?
  │
  ├── Yes → extract text/layout directly (PyMuPDF)
  │
  └── No  → rasterize page → PaddleOCR (local) → structured text
  │
  ▼
Merge all pages → rebuild as .docx (headings, paragraphs, tables)
  │
  ▼
Word doc out
```

Mixed PDFs (some pages scanned, some native) are handled page-by-page
automatically — you don't need to tell it which pages need OCR.

## Requirements

- Python 3.10+
- Packages: `pymupdf`, `python-docx`, `paddleocr`, `paddlepaddle`, `Pillow`,
  `opencv-python`, `pikepdf`
- No internet connection, API key, or account needed — everything runs on
  your machine

## Installation

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python convert.py input.pdf -o output.docx
```

Optional flags:

| Flag | Purpose |
|---|---|
| `--force-ocr` | Run OCR on every page even if a text layer exists (use for PDFs with a broken/garbled text layer) |
| `--lang` | OCR language, default `en` |
| `--dpi` | Rasterization DPI for scanned pages, default `300` |
| `--min-confidence` | Confidence threshold (0–1) below which a page is flagged in the log for manual review, default `0.6` |

## Packaging into a standalone executable

```bash
pip install pyinstaller
pyinstaller --onefile --name pdf2docx convert.py
```

This produces a single executable in `dist/` that runs without needing
Python installed — usable like any other command-line tool.

---

## Edge cases and how the tool handles them

| Edge case | Risk | Mitigation |
|---|---|---|
| **Scanned pages with skew/rotation** | OCR accuracy drops sharply past ~5° skew | Deskew pass (OpenCV) runs before OCR on every rasterized page |
| **Low-resolution scans (<150 DPI source)** | Garbled or missing characters | Upscale before OCR; flag pages below a confidence threshold in the output log rather than silently guessing |
| **Mixed text + scanned pages in one PDF** | Naive tools OCR everything or nothing | Per-page text-layer detection, handled independently per page |
| **Multi-column layouts (newspapers, academic papers)** | OCR reads across columns, scrambling sentence order | Layout-aware extraction (Azure Layout model, or PaddleOCR's layout detection) instead of raw left-to-right OCR |
| **Tables** | Most convertible tools flatten tables into run-on text | Table structure detection reconstructs actual Word tables, not text blocks |
| **Handwritten text** | Standard OCR is unreliable on handwriting; PaddleOCR is tuned for printed text | Detected and flagged in the output log as low-confidence rather than silently producing wrong text — treat as a draft to proofread manually |
| **Password-protected / encrypted PDFs** | Tool can't open the file at all | Detected upfront with a clear error rather than a silent crash; user must supply the password |
| **Corrupted or malformed PDFs** | Parser crashes mid-file | Wrapped in error handling; attempts a repair pass (via `pikepdf`) before giving up |
| **Non-English text / mixed-language documents** | Wrong language model = garbage output | `--lang` flag, with auto-detection fallback per page where possible |
| **Embedded images that aren't OCR targets (logos, photos, signatures)** | Tool tries to OCR meaningless image regions | Layout detection distinguishes body text regions from decorative/photo regions and skips the latter |
| **Very large PDFs (100+ pages)** | Memory issues, slow processing, API cost | Streamed page-by-page processing instead of loading the whole doc into memory; progress logging so it doesn't look hung |
| **PDFs with a text layer that's garbled or wrong (bad prior OCR baked in)** | Tool trusts the fake "native" text layer and skips OCR, producing garbage | `--force-ocr` flag overrides text-layer detection; tool also does a sanity check (gibberish-ratio heuristic) and can auto-flag suspect pages |
| **Footnotes, headers/footers, page numbers** | Get interleaved into body text at the wrong spot | Positional filtering keeps repeating header/footer regions separate from main content flow |
| **Forms with checkboxes/fields** | OCR reads checkbox glyphs as random characters | Selection-mark detection (available in both PaddleOCR layout models and Azure) instead of treating checkboxes as text |
| **Non-Latin scripts (Arabic, CJK, etc.)** | Wrong reading order (RTL) or missing character sets | PaddleOCR ships separate language models (`--lang ch`, `--lang arabic`, etc.) that must be selected explicitly — tool detects script mismatch and warns rather than silently mis-OCRing |
| **No GPU available** | OCR on CPU is noticeably slower, especially on large scanned batches | Runs on CPU by default with no extra setup; GPU auto-used if `paddlepaddle-gpu` + CUDA are detected, otherwise falls back cleanly |
| **Long-running batch job interrupted (crash, closed terminal)** | Have to restart from page 1 | Per-page progress checkpointing — a resumed run skips pages already converted |
| **Output docx formatting mismatch** | Bold/italic/headings guessed wrong from font metadata | Font-size and weight heuristics for headings; conservative fallback to plain paragraphs when confidence is low, rather than inventing structure that isn't there |

## Known limitations (being upfront)

- Handwriting accuracy will never be as reliable as typed text — treat OCR'd handwriting as a draft to proofread, not a final transcription.
- Extremely dense or unusual table layouts (merged cells, rotated headers) may still need manual cleanup.
- Local OCR engines generally trail cloud services on very messy/low-quality scans and complex tables — good enough for most personal documents, but not flawless on the hardest cases.
- CPU-only processing is slower than a GPU or cloud pipeline — fine for occasional use, less fine for batch-converting hundreds of pages at once.
- This tool won't preserve exact visual fidelity (fonts, precise spacing) — it preserves *content and structure*, which is what actually matters for an editable Word doc.
