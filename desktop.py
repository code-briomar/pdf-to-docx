#!/usr/bin/env python3
"""Desktop entry point: runs app.py's FastAPI server in a background thread
and shows it in a native window via pywebview. This is what gets built into
the .exe.
"""
import threading

import uvicorn
import webview

from app import app
from updater import check_for_update

HOST = "127.0.0.1"
PORT = 8756  # ponytail: fixed local port, fine for a single-instance desktop app


def _run_server():
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


def main():
    check_for_update()  # relaunches and exits here if a newer build was applied

    threading.Thread(target=_run_server, daemon=True).start()
    webview.create_window("PDF to DOCX", f"http://{HOST}:{PORT}", width=560, height=520)
    webview.start()


if __name__ == "__main__":
    main()
