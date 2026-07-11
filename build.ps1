# Builds desktop.py into a single-file Windows exe.
# Run from an activated venv (venv311\Scripts\activate) with requirements installed.

pyinstaller --onefile --windowed --name pdf_to_docx `
    --add-data "convert.py;." `
    desktop.py

# ponytail: signing step, uncomment once you have a code-signing cert.
# Needs signtool.exe (Windows SDK) on PATH.
#
# signtool sign /f "path\to\cert.pfx" /p "$env:CERT_PASSWORD" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 dist\pdf_to_docx.exe

Write-Host "Built: dist\pdf_to_docx.exe"
