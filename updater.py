"""Silent self-update: check a hosted JSON manifest, download and replace the
running exe if a newer version is published. No-ops when not running as a
frozen PyInstaller exe (i.e. during normal development).
"""
import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

from version import VERSION, UPDATE_MANIFEST_URL


def check_for_update():
    if not UPDATE_MANIFEST_URL or not getattr(sys, "frozen", False):
        return  # ponytail: dev runs (python app.py) never self-update

    try:
        with urllib.request.urlopen(UPDATE_MANIFEST_URL, timeout=5) as r:
            manifest = json.load(r)
    except Exception:
        return  # offline or manifest unreachable -- just keep running

    latest = manifest.get("version", "")
    url = manifest.get("url", "")
    if not latest or not url or latest == VERSION:
        return

    try:
        new_exe = Path(tempfile.gettempdir()) / f"pdf_to_docx_{latest}.exe"
        urllib.request.urlretrieve(url, new_exe)
    except Exception:
        return  # download failed -- keep running current version

    current_exe = Path(sys.executable)
    _relaunch_with_replacement(current_exe, new_exe)
    sys.exit(0)


def _relaunch_with_replacement(current_exe: Path, new_exe: Path):
    # Windows won't let us overwrite a running exe, so a tiny batch script
    # waits for this process to exit, swaps the file, then relaunches it.
    pid = os.getpid()
    bat = Path(tempfile.gettempdir()) / "pdf_to_docx_update.bat"
    bat.write_text(f"""@echo off
:wait
tasklist /fi "PID eq {pid}" | find "{pid}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)
move /y "{new_exe}" "{current_exe}" >nul
start "" "{current_exe}"
del "%~f0"
""")
    subprocess.Popen(
        ["cmd", "/c", str(bat)],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )
