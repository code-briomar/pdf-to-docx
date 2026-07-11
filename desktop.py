#!/usr/bin/env python3
"""Desktop entry point: runs app.py's FastAPI server in a background thread
and shows it in a native window via pywebview. This is what gets built into
the .exe.
"""
import base64
import threading
from pathlib import Path

import uvicorn
import webview

from app import app
from updater import check_for_update

HOST = "127.0.0.1"
PORT = 8756  # ponytail: fixed local port, fine for a single-instance desktop app


def _run_server():
    uvicorn.run(app, host=HOST, port=PORT, log_level="warning")


class Api:
    """Exposed to the page's JS as window.pywebview.api. A pywebview window
    has no browser download manager, so the usual <a download> blob trick
    silently does nothing -- saving a file requires a native dialog instead.
    """

    def save_docx(self, filename, base64_data):
        path = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG, save_filename=filename,
            file_types=("Word Document (*.docx)",),
        )
        if not path:
            return False  # user cancelled
        Path(path if isinstance(path, str) else path[0]).write_bytes(
            base64.b64decode(base64_data)
        )
        return True


def main():
    check_for_update()  # relaunches and exits here if a newer build was applied

    threading.Thread(target=_run_server, daemon=True).start()
    webview.create_window(
        "PDF to DOCX", f"http://{HOST}:{PORT}", width=560, height=520, js_api=Api()
    )
    webview.start()


if __name__ == "__main__":
    main()
