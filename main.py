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

# ── Palette ───────────────────────────────────────────────────────────────────
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

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self):
        # Bottom items packed first so they anchor correctly
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

        # Progress bar (anchored to bottom, hidden until conversion starts)
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

        # Drop zone
        zone = tk.Frame(
            self, bg=SURFACE, cursor="hand2",
            highlightbackground=BORDER, highlightthickness=1
        )
        zone.pack(fill="x", padx=24, pady=(18, 0))

        icon_lbl = tk.Label(
            zone, text="📄", font=("Segoe UI", 32),
            bg=SURFACE, fg=ACCENT, cursor="hand2"
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

        # Info row
        info = tk.Frame(self, bg=BG)
        info.pack(fill="x", padx=24, pady=(8, 0))

        self._path_lbl = tk.Label(
            info, text="", font=("Consolas", 8),
            bg=BG, fg=MUTED, anchor="w"
        )
        self._path_lbl.pack(fill="x")

        self._badge_lbl = tk.Label(
            info, text="", font=("Segoe UI", 10, "bold"),
            bg=BG, fg=ACCENT, anchor="w"
        )
        self._badge_lbl.pack(fill="x")

        # Convert button (packed after file is selected)
        self._btn = tk.Button(
            self, text="Convertir",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT, fg="white",
            activebackground=ACCENT2, activeforeground="white",
            bd=0, padx=32, pady=10, cursor="hand2",
            command=self._on_convert, relief="flat"
        )

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 520) // 2
        y = (self.winfo_screenheight() - 400) // 2
        self.geometry(f"520x400+{x}+{y}")

    # ── Event handlers ─────────────────────────────────────────────────────────

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
        self._path_lbl.config(
            text=name if len(name) <= 55 else f"...{name[-52:]}"
        )

        if direction:
            src, dst = direction
            label = f"{src.lstrip('.').upper()}  →  {dst.lstrip('.').upper()}"
            color = ACCENT if dst == ".docx" else "#f59e0b"
            self._badge_lbl.config(text=label, fg=color)
            self._btn.pack(pady=(12, 0))
            self._set_status("Listo para convertir", MUTED)
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
        self._pb.pack(fill="x", padx=24, pady=(8, 0), before=self._btn)
        self._pb.start(12)
        self._set_status("Procesando…", MUTED)

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
        self._set_status(f"✓  {Path(out_path).name}", SUCCESS)
        if messagebox.askyesno("¡Listo!", "Conversión completada.\n\n¿Abrir el archivo ahora?"):
            os.startfile(out_path)

    def _on_error(self, msg: str):
        self._converting = False
        self._pb.stop()
        self._pb.pack_forget()
        self._btn.config(state="normal", text="Convertir")
        short = msg[:120] + ("…" if len(msg) > 120 else "")
        self._set_status(f"Error: {short}", ERROR)
        messagebox.showerror("Error de conversión", msg)

    def _set_status(self, text: str, color: str):
        self._status_var.set(text)
        self._status_lbl.config(fg=color)


if __name__ == "__main__":
    app = FoliaApp()
    app.mainloop()
