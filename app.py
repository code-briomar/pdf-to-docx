#!/usr/bin/env python3
"""Minimal local web UI for convert.py: upload a PDF, wait, download the docx."""
import contextlib
import io
import sys
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, FileResponse

import convert as convert_module

JOBS_DIR = Path(tempfile.gettempdir()) / "pdf_to_docx_jobs"
JOBS_DIR.mkdir(exist_ok=True)

app = FastAPI()


@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <!doctype html>
    <html>
    <head>
    <title>PDF to DOCX</title>
    <style>
      :root {
        --color-ink: #292929;
        --color-paper: #ffffff;
        --font-nh: Inter, "Helvetica Neue", Arial, sans-serif;
        --font-s-condensed: "Roboto Condensed", "Barlow Condensed", "Univers Condensed", sans-serif;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        background: var(--color-paper);
        color: var(--color-ink);
        font-family: var(--font-nh);
      }
      .cell {
        border: 2px solid var(--color-ink);
        border-radius: 0;
        padding: 43px 45px;
      }
      h1 {
        font-family: var(--font-nh);
        font-weight: 300;
        font-size: 43px;
        line-height: 1.34;
        letter-spacing: -0.02em;
        margin: 0 0 8px;
      }
      .label {
        font-family: var(--font-s-condensed);
        font-weight: 500;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        margin: 0 0 20px;
      }
      form {
        margin-top: 20px;
        display: flex;
        flex-direction: column;
        gap: 20px;
        max-width: 480px;
      }
      input[type="file"] {
        position: absolute;
        width: 1px;
        height: 1px;
        overflow: hidden;
        clip: rect(0, 0, 0, 0);
      }
      .file-field {
        display: flex;
        align-items: baseline;
        gap: 20px;
        border-bottom: 1px solid #000000;
        padding: 8px 0;
      }
      .file-field label {
        background: transparent;
        border: 1px solid var(--color-ink);
        border-radius: 0;
        padding: 8px 20px;
        color: var(--color-ink);
        font-family: var(--font-s-condensed);
        font-weight: 500;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        cursor: pointer;
        flex-shrink: 0;
      }
      .file-field label:hover { background: var(--color-ink); color: var(--color-paper); }
      .file-field .file-name {
        font-family: var(--font-nh);
        font-size: 16px;
        color: #000000;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      button {
        background: transparent;
        border: 1px solid var(--color-ink);
        border-radius: 0;
        padding: 12px 20px;
        color: var(--color-ink);
        font-family: var(--font-s-condensed);
        font-weight: 500;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        cursor: pointer;
        width: fit-content;
      }
      button:hover { background: var(--color-ink); color: var(--color-paper); }
      .caption {
        font-family: var(--font-s-condensed);
        font-weight: 300;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.2em;
        margin-top: 20px;
      }
      footer {
        border-top: 2px solid var(--color-ink);
        padding: 20px 45px;
        font-family: var(--font-s-condensed);
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.2em;
      }
      .clock {
        display: none;
        align-items: baseline;
        gap: 12px;
        border: 1px solid var(--color-ink);
        padding: 12px 20px;
        width: fit-content;
        margin-top: 20px;
      }
      .clock.running { display: flex; }
      .clock .ticks {
        font-family: var(--font-nh);
        font-weight: 100;
        font-size: 32px;
        letter-spacing: -0.02em;
        font-variant-numeric: tabular-nums;
      }
      .clock .ticks-label {
        font-family: var(--font-s-condensed);
        font-weight: 500;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
      }
      .error {
        display: none;
        border: 1px solid var(--color-ink);
        padding: 12px 20px;
        margin-top: 20px;
        font-family: var(--font-s-condensed);
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.1em;
      }
      .error.visible { display: block; }
    </style>
    </head>
    <body>
      <div class="cell">
        <p class="label">Local Conversion Tool</p>
        <h1>PDF to DOCX</h1>
        <form id="convert-form">
          <div class="file-field">
            <label for="file-input">Choose File</label>
            <span class="file-name" id="file-name">No file chosen</span>
          </div>
          <input type="file" id="file-input" name="file" accept=".pdf" required>
          <button type="submit" id="convert-button">Convert</button>
        </form>
        <div class="clock" id="clock">
          <span class="ticks" id="ticks">00:00</span>
          <span class="ticks-label" id="ticks-label">Converting</span>
        </div>
        <div class="error" id="error"></div>
        <p class="caption">Conversion runs synchronously. Large scanned PDFs may take a while.</p>
      </div>
      <footer>&copy; Local Instance</footer>
      <script>
        const fileInput = document.getElementById('file-input');
        const fileName = document.getElementById('file-name');
        fileInput.addEventListener('change', () => {
          fileName.textContent = fileInput.files.length ? fileInput.files[0].name : 'No file chosen';
        });

        const form = document.getElementById('convert-form');
        const button = document.getElementById('convert-button');
        const clock = document.getElementById('clock');
        const ticks = document.getElementById('ticks');
        const ticksLabel = document.getElementById('ticks-label');
        const errorBox = document.getElementById('error');

        form.addEventListener('submit', async (e) => {
          e.preventDefault();
          if (!fileInput.files.length) return;

          errorBox.classList.remove('visible');
          button.disabled = true;
          clock.classList.add('running');
          ticksLabel.textContent = 'Converting';
          const start = Date.now();
          const timer = setInterval(() => {
            const elapsed = Math.floor((Date.now() - start) / 1000);
            const m = String(Math.floor(elapsed / 60)).padStart(2, '0');
            const s = String(elapsed % 60).padStart(2, '0');
            ticks.textContent = m + ':' + s;
          }, 250);

          try {
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            const res = await fetch('/convert', { method: 'POST', body: formData });
            clearInterval(timer);
            if (!res.ok) {
              const detail = await res.json().catch(() => ({}));
              throw new Error(detail.detail || ('Conversion failed (HTTP ' + res.status + ')'));
            }
            ticksLabel.textContent = 'Done';
            const disposition = res.headers.get('Content-Disposition') || '';
            const match = disposition.match(/filename="?([^"]+)"?/);
            const outName = match ? match[1] : 'output.docx';
            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = outName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
          } catch (err) {
            clearInterval(timer);
            ticksLabel.textContent = 'Failed';
            errorBox.textContent = err.message;
            errorBox.classList.add('visible');
          } finally {
            button.disabled = false;
          }
        });
      </script>
    </body>
    </html>
    """


@app.post("/convert")
def convert(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Please upload a .pdf file")

    job_id = uuid.uuid4().hex
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir()
    input_path = job_dir / "input.pdf"
    output_path = job_dir / "output.docx"

    input_path.write_bytes(file.file.read())

    # ponytail: in-process call, not a subprocess -- a frozen desktop exe has no
    # separate python interpreter for sys.executable to point at. sys.argv swap
    # is not thread-safe under concurrent requests; fine for single-user desktop use.
    stderr = io.StringIO()
    old_argv = sys.argv
    sys.argv = ["convert.py", str(input_path), "-o", str(output_path)]
    try:
        with contextlib.redirect_stderr(stderr), contextlib.redirect_stdout(stderr):
            convert_module.main()
        error = None
    except SystemExit as e:
        error = f"exit code {e.code}" if e.code else None
    except Exception as e:
        error = str(e)
    finally:
        sys.argv = old_argv

    if error or not output_path.exists():
        raise HTTPException(500, f"Conversion failed: {error or stderr.getvalue()[-2000:]}")

    out_name = Path(file.filename).stem + ".docx"
    return FileResponse(
        output_path, filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
