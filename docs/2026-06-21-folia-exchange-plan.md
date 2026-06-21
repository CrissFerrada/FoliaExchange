# FoliaExchange Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build FoliaExchange, un convertidor PDF↔DOCX standalone para Windows como EXE hermoso sin depender de Word ni LibreOffice.

**Architecture:** Single tkinter window with a dark premium UI. Converter functions live in `converters/` as independent modules, registered in a dict keyed by `(src_ext, dst_ext)` inside `main.py`. GUI picks converter by file extension, opens a Save As dialog, runs conversion in a background thread, shows progress bar.

**Tech Stack:** Python 3.11+, tkinter (stdlib), pdf2docx, mammoth, weasyprint, PyInstaller 6+, pytest, python-docx (test only)

## Global Constraints

- Project root: `C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange\`
- No Word, no LibreOffice dependency — must run on any Windows 11 PC
- Credits line verbatim: `Hecho por Cristofher Ferrada`
- App name verbatim: `FoliaExchange`
- Color palette: bg `#0f0f1a`, surface `#1a1a2e`, accent `#6366f1`, text `#e2e8f0`, muted `#64748b`
- Font: Segoe UI throughout
- GitHub account: Cristofher
- Obsidian vault repo: `obsidian-vault` (already exists at github.com/Cristofher/obsidian-vault)

---

## File Map

| File | Responsibility |
|---|---|
| `main.py` | GUI + CONVERTERS registry + event handlers |
| `converters/pdf_to_docx.py` | pdf2docx wrapper — `convert(in, out)` |
| `converters/docx_to_pdf.py` | mammoth→weasyprint pipeline — `convert(in, out)` |
| `converters/__init__.py` | Empty (namespace package) |
| `tests/conftest.py` | pytest fixtures: `sample_pdf`, `sample_docx` |
| `tests/test_pdf_to_docx.py` | Unit tests for pdf_to_docx.convert |
| `tests/test_docx_to_pdf.py` | Unit tests for docx_to_pdf.convert |
| `requirements.txt` | Production deps |
| `requirements-dev.txt` | Test deps (extends requirements.txt) |
| `build.bat` | Install deps + PyInstaller → FoliaExchange.exe |
| `update.bat` | git pull + rebuild |
| `.gitignore` | Excludes dist/, build/, __pycache__, *.spec |
| `README.md` | Project description, usage, credits |

---

## Task 1: Project scaffold + PDF → DOCX converter

**Files:**
- Create: `converters/__init__.py`
- Create: `converters/pdf_to_docx.py`
- Create: `requirements.txt`
- Create: `requirements-dev.txt`
- Create: `tests/conftest.py`
- Test: `tests/test_pdf_to_docx.py`

**Interfaces:**
- Produces: `converters.pdf_to_docx.convert(input_path: str, output_path: str) -> None` — raises `Exception` on failure

- [ ] **Step 1: Create project directory structure**

```bash
mkdir "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange"
cd "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange"
mkdir converters tests
```

- [ ] **Step 2: Create `requirements.txt`**

```
pdf2docx
mammoth
weasyprint
pyinstaller
```

- [ ] **Step 3: Create `requirements-dev.txt`**

```
-r requirements.txt
pytest
python-docx
```

- [ ] **Step 4: Install dev dependencies**

```bash
pip install -r requirements-dev.txt
```

Expected: all packages install without errors.

- [ ] **Step 5: Create `converters/__init__.py`** (empty namespace package)

```python
```

- [ ] **Step 6: Write the failing test**

`tests/conftest.py`:
```python
import pytest
import weasyprint
from docx import Document


@pytest.fixture
def sample_pdf(tmp_path):
    path = tmp_path / "sample.pdf"
    weasyprint.HTML(string="<p>Hello FoliaExchange test.</p>").write_pdf(str(path))
    return str(path)


@pytest.fixture
def sample_docx(tmp_path):
    doc = Document()
    doc.add_heading("FoliaExchange Test", 0)
    doc.add_paragraph("Hello, this is a test document.")
    path = tmp_path / "sample.docx"
    doc.save(str(path))
    return str(path)
```

`tests/test_pdf_to_docx.py`:
```python
import pytest
from pathlib import Path
from converters.pdf_to_docx import convert


def test_creates_docx_file(sample_pdf, tmp_path):
    out = str(tmp_path / "output.docx")
    convert(sample_pdf, out)
    assert Path(out).exists()
    assert Path(out).stat().st_size > 1000


def test_raises_on_missing_file(tmp_path):
    with pytest.raises(Exception):
        convert(str(tmp_path / "ghost.pdf"), str(tmp_path / "out.docx"))
```

- [ ] **Step 7: Run test — verify it FAILS**

```bash
cd "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange"
python -m pytest tests/test_pdf_to_docx.py -v
```

Expected: `ImportError` or `ModuleNotFoundError` — `converters.pdf_to_docx` doesn't exist yet.

- [ ] **Step 8: Create `converters/pdf_to_docx.py`**

```python
from pdf2docx import Converter


def convert(input_path: str, output_path: str) -> None:
    """Convert PDF to DOCX using pdf2docx. Raises on failure."""
    cv = Converter(input_path)
    try:
        cv.convert(output_path, start=0, end=None)
    finally:
        cv.close()
```

- [ ] **Step 9: Run test — verify it PASSES**

```bash
python -m pytest tests/test_pdf_to_docx.py -v
```

Expected:
```
PASSED tests/test_pdf_to_docx.py::test_creates_docx_file
PASSED tests/test_pdf_to_docx.py::test_raises_on_missing_file
```

- [ ] **Step 10: Commit**

```bash
git init
git add converters/__init__.py converters/pdf_to_docx.py requirements.txt requirements-dev.txt tests/conftest.py tests/test_pdf_to_docx.py
git commit -m "feat: scaffold + PDF→DOCX converter"
```

---

## Task 2: DOCX → PDF converter

**Files:**
- Create: `converters/docx_to_pdf.py`
- Test: `tests/test_docx_to_pdf.py`

**Interfaces:**
- Consumes: `sample_docx` fixture from `tests/conftest.py`
- Produces: `converters.docx_to_pdf.convert(input_path: str, output_path: str) -> None` — raises `Exception` on failure

- [ ] **Step 1: Write the failing test**

`tests/test_docx_to_pdf.py`:
```python
import pytest
from pathlib import Path
from converters.docx_to_pdf import convert


def test_creates_pdf_file(sample_docx, tmp_path):
    out = str(tmp_path / "output.pdf")
    convert(sample_docx, out)
    assert Path(out).exists()
    assert Path(out).stat().st_size > 1000
    with open(out, "rb") as f:
        assert f.read(4) == b"%PDF"


def test_raises_on_missing_file(tmp_path):
    with pytest.raises(Exception):
        convert(str(tmp_path / "ghost.docx"), str(tmp_path / "out.pdf"))
```

- [ ] **Step 2: Run test — verify it FAILS**

```bash
python -m pytest tests/test_docx_to_pdf.py -v
```

Expected: `ImportError` — `converters.docx_to_pdf` doesn't exist yet.

- [ ] **Step 3: Create `converters/docx_to_pdf.py`**

```python
import mammoth
import weasyprint

_CSS = """
body {
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    margin: 2.54cm;
    color: #000;
}
h1 { font-size: 16pt; margin-bottom: 6pt; }
h2 { font-size: 13pt; margin-bottom: 4pt; }
h3 { font-size: 12pt; margin-bottom: 4pt; }
p  { margin: 0 0 6pt; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; }
td, th { border: 1px solid #bbb; padding: 4pt 6pt; }
th { background: #f0f0f0; font-weight: bold; }
img { max-width: 100%; }
"""


def convert(input_path: str, output_path: str) -> None:
    """Convert DOCX to PDF via mammoth→HTML→weasyprint. Raises on failure."""
    with open(input_path, "rb") as f:
        result = mammoth.convert_to_html(f)
    html = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        f"<style>{_CSS}</style></head>"
        f"<body>{result.value}</body></html>"
    )
    weasyprint.HTML(string=html).write_pdf(output_path)
```

- [ ] **Step 4: Run test — verify it PASSES**

```bash
python -m pytest tests/test_docx_to_pdf.py -v
```

Expected:
```
PASSED tests/test_docx_to_pdf.py::test_creates_pdf_file
PASSED tests/test_docx_to_pdf.py::test_raises_on_missing_file
```

- [ ] **Step 5: Run full suite**

```bash
python -m pytest tests/ -v
```

Expected: 4 tests, all PASSED.

- [ ] **Step 6: Commit**

```bash
git add converters/docx_to_pdf.py tests/test_docx_to_pdf.py
git commit -m "feat: DOCX→PDF converter via mammoth + weasyprint"
```

---

## Task 3: GUI — main.py

**Files:**
- Create: `main.py`

**Interfaces:**
- Consumes: `converters.pdf_to_docx.convert` and `converters.docx_to_pdf.convert`
- Produces: runnable `FoliaApp` tkinter window

- [ ] **Step 1: Create `main.py`**

```python
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from pathlib import Path

from converters.pdf_to_docx import convert as _pdf_to_docx
from converters.docx_to_pdf import convert as _docx_to_pdf

CONVERTERS = {
    (".pdf",  ".docx"): _pdf_to_docx,
    (".docx", ".pdf"):  _docx_to_pdf,
    (".doc",  ".pdf"):  _docx_to_pdf,
}

# ── Palette ──────────────────────────────────────────────────────────────────
BG      = "#0f0f1a"
SURFACE = "#1a1a2e"
ACCENT  = "#6366f1"
ACCENT2 = "#4f46e5"
TEXT    = "#e2e8f0"
MUTED   = "#64748b"
SUCCESS = "#10b981"
ERROR   = "#ef4444"
BORDER  = "#2d2d4e"
HDR_BG  = "#0a0a15"


class FoliaApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FoliaExchange")
        self.geometry("520x400")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._filepath = None
        self._converting = False
        self._build()
        self._center()

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build(self):
        # Bottom items first so they anchor correctly
        tk.Label(
            self, text="Hecho por Cristofher Ferrada",
            font=("Segoe UI", 8), bg=BG, fg=MUTED
        ).pack(side="bottom", pady=(0, 6))

        self._status_var = tk.StringVar()
        self._status_lbl = tk.Label(
            self, textvariable=self._status_var,
            font=("Segoe UI", 9), bg=BG, fg=MUTED
        )
        self._status_lbl.pack(side="bottom", pady=(0, 2))

        # Progress bar (always in layout, starts hidden via zero height trick)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "FE.Horizontal.TProgressbar",
            troughcolor=SURFACE, background=ACCENT,
            borderwidth=0, thickness=5
        )
        self._pb = ttk.Progressbar(
            self, style="FE.Horizontal.TProgressbar", mode="indeterminate"
        )
        self._pb.pack(fill="x", padx=24, pady=(0, 4), side="bottom")
        self._pb.pack_forget()  # hidden until conversion starts

        # Header
        hdr = tk.Frame(self, bg=HDR_BG, pady=14)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="FoliaExchange",
            font=("Segoe UI", 22, "bold"), bg=HDR_BG, fg=TEXT
        ).pack()
        tk.Label(
            hdr, text="Convierte documentos sin complicaciones",
            font=("Segoe UI", 9), bg=HDR_BG, fg=MUTED
        ).pack(pady=(2, 0))

        # Drop zone (clickable)
        zone = tk.Frame(
            self, bg=SURFACE, cursor="hand2",
            highlightbackground=BORDER, highlightthickness=1
        )
        zone.pack(fill="x", padx=24, pady=(18, 0))

        icon_lbl = tk.Label(
            zone, text="📄", font=("Segoe UI", 32), bg=SURFACE, fg=ACCENT, cursor="hand2"
        )
        icon_lbl.pack(pady=(14, 2))

        hint_lbl = tk.Label(
            zone, text="Haz clic para seleccionar archivo",
            font=("Segoe UI", 11, "bold"), bg=SURFACE, fg=TEXT, cursor="hand2"
        )
        hint_lbl.pack()

        fmt_lbl = tk.Label(
            zone, text="PDF  ·  DOCX  ·  DOC",
            font=("Segoe UI", 9), bg=SURFACE, fg=MUTED, cursor="hand2"
        )
        fmt_lbl.pack(pady=(2, 14))

        for w in (zone, icon_lbl, hint_lbl, fmt_lbl):
            w.bind("<Button-1>", lambda e: self._browse())

        # Info section (path + direction badge)
        info = tk.Frame(self, bg=BG)
        info.pack(fill="x", padx=24, pady=(8, 0))

        self._path_lbl = tk.Label(
            info, text="", font=("Consolas", 8), bg=BG, fg=MUTED, anchor="w"
        )
        self._path_lbl.pack(fill="x")

        self._badge_lbl = tk.Label(
            info, text="", font=("Segoe UI", 10, "bold"), bg=BG, fg=ACCENT, anchor="w"
        )
        self._badge_lbl.pack(fill="x")

        # Convert button (hidden until file selected)
        self._btn = tk.Button(
            self, text="Convertir",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT, fg="white",
            activebackground=ACCENT2, activeforeground="white",
            bd=0, padx=32, pady=10, cursor="hand2",
            command=self._on_convert, relief="flat"
        )
        # Not packed yet — shown in _set_file

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 520) // 2
        y = (self.winfo_screenheight() - 400) // 2
        self.geometry(f"520x400+{x}+{y}")

    # ── Event handlers ────────────────────────────────────────────────────────

    def _browse(self):
        if self._converting:
            return
        path = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[
                ("Documentos", "*.pdf *.docx *.doc"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx *.doc"),
                ("Todos los archivos", "*.*"),
            ]
        )
        if path:
            self._set_file(path)

    def _set_file(self, path: str):
        self._filepath = path
        ext = Path(path).suffix.lower()

        direction = next(((s, d) for s, d in CONVERTERS if s == ext), None)

        name = Path(path).name
        self._path_lbl.config(text=name if len(name) <= 55 else f"...{name[-52:]}")

        if direction:
            src, dst = direction
            label = f"{src.lstrip('.').upper()}  →  {dst.lstrip('.').upper()}"
            color = ACCENT if dst == ".docx" else "#f59e0b"
            self._badge_lbl.config(text=label, fg=color)
            self._btn.pack(pady=(12, 0))
            self._status_var.set("Listo para convertir")
            self._status_lbl.config(fg=MUTED)
        else:
            self._badge_lbl.config(text="Formato no soportado", fg=ERROR)
            self._btn.pack_forget()

        self._pb.pack_forget()

    def _on_convert(self):
        if not self._filepath or self._converting:
            return

        ext = Path(self._filepath).suffix.lower()
        direction = next(((s, d) for s, d in CONVERTERS if s == ext), None)
        if not direction:
            return

        _, dst = direction
        default_name = Path(self._filepath).stem + dst
        out_path = filedialog.asksaveasfilename(
            title="Guardar archivo convertido",
            initialfile=default_name,
            initialdir=str(Path(self._filepath).parent),
            defaultextension=dst,
            filetypes=[(f"Archivo {dst.lstrip('.').upper()}", f"*{dst}")]
        )
        if not out_path:
            return

        self._converting = True
        self._btn.config(state="disabled", text="Convirtiendo...")
        self._pb.pack(fill="x", padx=24, pady=(8, 0))
        self._pb.start(12)
        self._status_var.set("Procesando…")
        self._status_lbl.config(fg=MUTED)

        def worker():
            try:
                fn = CONVERTERS[(ext, dst)]
                fn(self._filepath, out_path)
                self.after(0, self._on_success, out_path)
            except Exception as exc:
                self.after(0, self._on_error, str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, out_path: str):
        self._converting = False
        self._pb.stop()
        self._pb.pack_forget()
        self._btn.config(state="normal", text="Convertir")
        self._status_var.set(f"✓  {Path(out_path).name}")
        self._status_lbl.config(fg=SUCCESS)
        if messagebox.askyesno("¡Listo!", "Conversión completada.\n\n¿Abrir el archivo ahora?"):
            os.startfile(out_path)

    def _on_error(self, msg: str):
        self._converting = False
        self._pb.stop()
        self._pb.pack_forget()
        self._btn.config(state="normal", text="Convertir")
        short = msg[:120] + ("…" if len(msg) > 120 else "")
        self._status_var.set(f"Error: {short}")
        self._status_lbl.config(fg=ERROR)
        messagebox.showerror("Error de conversión", msg)


if __name__ == "__main__":
    app = FoliaApp()
    app.mainloop()
```

- [ ] **Step 2: Run the app manually and verify UI**

```bash
python main.py
```

Verify:
- Window opens centered, 520×400 px
- Dark background, indigo accent visible
- Click zone opens file dialog
- After selecting a PDF: badge shows `PDF → DOCX`, Convert button appears
- After selecting a DOCX: badge shows `DOCX → PDF`, badge is amber
- Footer shows `Hecho por Cristofher Ferrada`
- Conversion runs (progress bar animates), Save As dialog opens, output file created
- Success message asks to open file

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: GUI — FoliaApp tkinter window"
```

---

## Task 4: Build scripts + .gitignore

**Files:**
- Create: `build.bat`
- Create: `update.bat`
- Create: `.gitignore`

- [ ] **Step 1: Create `.gitignore`**

```
dist/
build/
*.spec
__pycache__/
*.pyc
*.pyo
.pytest_cache/
*.egg-info/
.env
```

- [ ] **Step 2: Create `build.bat`**

```bat
@echo off
echo.
echo  ====================================
echo   FoliaExchange — Build
echo  ====================================
echo.
echo [1/2] Instalando dependencias...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo ERROR: fallo al instalar dependencias.
    pause
    exit /b 1
)
echo.
echo [2/2] Construyendo FoliaExchange.exe...
pyinstaller --onefile --windowed --name FoliaExchange ^
  --collect-all pdf2docx ^
  --collect-all weasyprint ^
  --collect-all mammoth ^
  --hidden-import PIL ^
  main.py
if %errorlevel% neq 0 (
    echo ERROR: fallo en PyInstaller.
    pause
    exit /b 1
)
echo.
echo  ====================================
echo   Listo! dist\FoliaExchange.exe
echo  ====================================
echo.
pause
```

- [ ] **Step 3: Create `update.bat`**

```bat
@echo off
echo.
echo  ====================================
echo   FoliaExchange — Actualizar
echo  ====================================
echo.
echo [1/2] Descargando ultima version...
git pull origin main
if %errorlevel% neq 0 (
    echo ERROR: fallo en git pull.
    pause
    exit /b 1
)
echo.
echo [2/2] Reconstruyendo ejecutable...
call build.bat
```

- [ ] **Step 4: Verify build works**

Double-click `build.bat` or run:
```bash
build.bat
```

Expected: `dist\FoliaExchange.exe` appears (~80 MB). Double-click it and confirm the app opens without Python installed.

- [ ] **Step 5: Commit**

```bash
git add .gitignore build.bat update.bat
git commit -m "feat: build + update scripts"
```

---

## Task 5: README + GitHub repository

**Files:**
- Create: `README.md`
- Actions: create GitHub repo, push, tag initial release

- [ ] **Step 1: Create `README.md`**

```markdown
# FoliaExchange

Convierte entre PDF y Word (DOCX) de forma rápida y sin instalar Word ni LibreOffice.

## Características

- **PDF → DOCX** — preserva texto, tablas e imágenes
- **DOCX → PDF** — fiel al contenido original
- Interfaz oscura y moderna
- Diálogo "Guardar como" para elegir nombre y carpeta
- Sin dependencias externas — todo en un solo `.exe`

## Uso

1. Abre `FoliaExchange.exe`
2. Haz clic en la zona de selección y elige tu archivo (PDF, DOCX o DOC)
3. Haz clic en **Convertir**
4. Elige dónde guardar el resultado (puedes renombrarlo)
5. ¡Listo!

## Formatos soportados

| Entrada | Salida |
|---|---|
| PDF | DOCX |
| DOCX | PDF |
| DOC | PDF |

## Construir desde código fuente

Requiere Python 3.11+.

```bash
pip install -r requirements-dev.txt
python main.py          # ejecutar en desarrollo
build.bat               # generar FoliaExchange.exe
update.bat              # actualizar y reconstruir
```

## Ejecutar tests

```bash
python -m pytest tests/ -v
```

---

Hecho por **Cristofher Ferrada**
```

- [ ] **Step 2: Create the GitHub repository**

```bash
cd "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange"
gh repo create FoliaExchange --public --description "Convertidor PDF↔DOCX standalone para Windows. Sin Word, sin LibreOffice." --source . --remote origin --push
```

Expected: repo created at `https://github.com/Cristofher/FoliaExchange` and all commits pushed.

- [ ] **Step 3: Verify push**

```bash
gh repo view Cristofher/FoliaExchange
```

Expected: repo info displayed with correct description.

- [ ] **Step 4: Commit README**

```bash
git add README.md
git commit -m "docs: README with usage and build instructions"
git push origin main
```

---

## Task 6: Obsidian vault note

**Files:**
- Create: `Proyectos/FoliaExchange.md` inside the `obsidian-vault` repo

- [ ] **Step 1: Clone obsidian-vault (if not already cloned)**

```bash
cd "C:\Users\crist\OneDrive\Desktop\Proyectos"
gh repo clone Cristofher/obsidian-vault
```

If already cloned, pull latest:
```bash
cd "C:\...\obsidian-vault"
git pull origin main
```

- [ ] **Step 2: Verify the Proyectos folder exists**

```bash
ls "C:\Users\crist\OneDrive\Desktop\Proyectos\obsidian-vault\Proyectos"
```

If folder doesn't exist:
```bash
mkdir "C:\Users\crist\OneDrive\Desktop\Proyectos\obsidian-vault\Proyectos"
```

- [ ] **Step 3: Create `obsidian-vault/Proyectos/FoliaExchange.md`**

```markdown
---
tags: [proyecto, herramienta, python, utilidad]
fecha: 2026-06-21
estado: activo
repo: https://github.com/Cristofher/FoliaExchange
---

# FoliaExchange

Aplicación de escritorio para Windows que convierte documentos entre **PDF** y **Word (DOCX)** sin necesitar Word ni LibreOffice.

## Stack técnico

- Python 3.11 + tkinter
- `pdf2docx` — conversión PDF → DOCX
- `mammoth` + `weasyprint` — conversión DOCX → PDF
- PyInstaller — genera un `.exe` standalone

## Formatos soportados

- PDF → DOCX
- DOCX → PDF
- DOC → PDF

## Cómo usar

1. Descargar `FoliaExchange.exe` desde el repo
2. Doble clic — no requiere instalación
3. Seleccionar archivo → Convertir → Guardar como

## Cómo actualizar / reconstruir

```bash
update.bat   # git pull + rebuild
build.bat    # solo rebuild
```

## Extensiones futuras planeadas

- MD → PDF
- PDF → MD
- Conversión por lotes

## Hecho por

Cristofher Ferrada — [[Perfil]]
```

- [ ] **Step 4: Commit and push to obsidian-vault**

```bash
cd "C:\Users\crist\OneDrive\Desktop\Proyectos\obsidian-vault"
git add Proyectos/FoliaExchange.md
git commit -m "docs: add FoliaExchange project note"
git push origin main
```

Expected: note appears in GitHub at `obsidian-vault/Proyectos/FoliaExchange.md`.

---

## Task 7: Mover docs de diseño fuera de My Mineral Warriors

Los archivos de diseño se crearon dentro del proyecto My Mineral Warriors por conveniencia. Este paso los mueve a la carpeta correcta de FoliaExchange.

**Files:**
- Move: `My Mineral Warriors/docs/superpowers/specs/2026-06-21-folia-exchange-design.md` → `FoliaExchange/docs/2026-06-21-folia-exchange-design.md`
- Move: `My Mineral Warriors/docs/superpowers/plans/2026-06-21-folia-exchange.md` → `FoliaExchange/docs/2026-06-21-folia-exchange-plan.md`

- [ ] **Step 1: Create docs folder in FoliaExchange**

```bash
mkdir "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange\docs"
```

- [ ] **Step 2: Move spec and plan files**

```bash
move "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\My Mineral Warriors\docs\superpowers\specs\2026-06-21-folia-exchange-design.md" "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange\docs\2026-06-21-folia-exchange-design.md"

move "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\My Mineral Warriors\docs\superpowers\plans\2026-06-21-folia-exchange.md" "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange\docs\2026-06-21-folia-exchange-plan.md"
```

- [ ] **Step 3: Commit the moved docs to FoliaExchange**

```bash
cd "C:\Users\crist\OneDrive\Desktop\Proyectos\Aplicaciones\FoliaExchange"
git add docs/
git commit -m "docs: add design spec and implementation plan"
git push origin main
```
