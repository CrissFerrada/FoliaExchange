# FoliaExchange

Convierte entre PDF y Word (DOCX) de forma rápida y sin instalar Word ni LibreOffice.

## Características

- **PDF → DOCX** — documento limpio y editable: detecta columnas (papers), conserva títulos e imágenes y reacomoda el texto en párrafos normales (sin hojas en blanco)
- **DOCX → PDF** — fiel al contenido original
- Interfaz oscura y moderna
- Diálogo "Guardar como" para elegir nombre y carpeta de destino
- Sin dependencias externas — todo en un solo `.exe`

## Uso

1. Abre `FoliaExchange.exe`
2. Haz clic en la zona de selección y elige tu archivo (PDF, DOCX o DOC)
3. Haz clic en **Convertir**
4. Elige dónde guardar el resultado (puedes renombrarlo)
5. ¡Listo!

## Formatos soportados

| Entrada | Salida |
|---------|--------|
| PDF     | DOCX   |
| DOCX    | PDF    |
| DOC     | PDF    |

## Construir desde código fuente

Requiere Python 3.11+.

```bash
pip install -r requirements-dev.txt
python main.py      # ejecutar en desarrollo
build.bat           # generar FoliaExchange.exe en dist/
update.bat          # git pull + rebuild
```

## Ejecutar tests

```bash
python -m pytest tests/ -v
```

---

Hecho por **Cristofher Ferrada**
