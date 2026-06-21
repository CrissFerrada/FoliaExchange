# FoliaExchange — Design Spec
**Date:** 2026-06-21  
**Author:** Cristofher Ferrada  
**Status:** Approved

---

## Overview

FoliaExchange is a standalone Windows desktop application that converts documents between PDF and DOCX formats. It requires no external software (no Word, no LibreOffice). Distributed as a single `FoliaExchange.exe`.

---

## Identity

- **App name:** FoliaExchange
- **Tagline:** Convierte documentos sin complicaciones
- **Credits line:** Hecho por Cristofher Ferrada
- **Color palette:** Dark background `#0f0f1a` · Indigo accent `#6366f1` · Cream text `#e2e8f0`
- **Font:** Segoe UI (native Windows)

---

## Architecture

```
FoliaExchange/
  main.py                  — GUI (tkinter) + converter orchestrator
  requirements.txt         — pdf2docx, mammoth, weasyprint, pyinstaller
  converters/
    __init__.py
    pdf_to_docx.py         — pdf2docx wrapper
    docx_to_pdf.py         — mammoth → weasyprint pipeline
  build.bat                — installs deps + builds FoliaExchange.exe
  update.bat               — git pull + rebuild
  .gitignore
  README.md
```

### Converter registry (extensible)

```python
CONVERTERS = {
    (".pdf",  ".docx"): converters.pdf_to_docx,
    (".docx", ".pdf"):  converters.docx_to_pdf,
    (".doc",  ".pdf"):  converters.docx_to_pdf,
    # Future:
    # (".md",  ".pdf"): converters.md_to_pdf,
}
```

Adding a new format = one new file + one new dict entry. The GUI requires no changes.

---

## Conversion Stack

| Direction | Library | Notes |
|---|---|---|
| PDF → DOCX | `pdf2docx` | Preserves text, tables, images |
| DOCX → HTML | `mammoth` | Faithful structural conversion |
| HTML → PDF | `weasyprint` | Uses system fonts (Calibri, Arial, etc.) |
| EXE bundling | `PyInstaller --onefile` | ~80 MB, fully standalone |

---

## User Flow

### Main screen
1. User opens `FoliaExchange.exe`
2. Centered window (520×380 px) shows:
   - App name + tagline at top
   - Large clickable zone "Seleccionar archivo" with page icon
   - Supported formats shown: `PDF · DOCX · DOC`
   - "Hecho por Cristofher Ferrada" in footer
3. No file selected → Convert button is hidden

### After file selection
1. File path shown below the selection zone
2. Auto-detected badge appears: `PDF → DOCX` or `DOCX → PDF`
3. "Convertir" button appears

### Conversion
1. User clicks "Convertir"
2. **Save As dialog** opens pre-filled: same filename, target extension
3. User can rename and choose folder, then confirms
4. Progress bar animates while conversion runs in background thread (non-blocking)
5. On success: toast/message "¡Listo!" + button "Abrir archivo"
6. On error: error message displayed in-app (no crash dialogs)

---

## Scripts

### `build.bat`
```bat
pip install -r requirements.txt
pyinstaller --onefile --windowed --name FoliaExchange main.py
```

### `update.bat`
```bat
git pull origin main
pip install -r requirements.txt
call build.bat
```

---

## GitHub Delivery

- New public repo: `FoliaExchange` under Cristofher's account
- `README.md` with: description, screenshot placeholder, usage instructions, credits
- `.gitignore` excludes: `dist/`, `build/`, `*.spec`, `__pycache__/`
- Initial commit includes all source, scripts, and README

---

## Obsidian Delivery

- File: `obsidian-vault/Proyectos/FoliaExchange.md`
- Contents: project description, repo link, how to build/run, formats supported, date created
- Push to `obsidian-vault` repo on GitHub

---

## Out of Scope (v1)

- Drag and drop file input
- Batch conversion (multiple files)
- Dark/light mode toggle
- MD → PDF (planned for v2, architecture already supports it)
- Scanned PDF / OCR support
