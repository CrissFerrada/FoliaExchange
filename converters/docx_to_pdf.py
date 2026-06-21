import mammoth
from xhtml2pdf import pisa
from io import BytesIO


_CSS = """
@page { margin: 2.54cm; }
body {
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #000;
}
h1 { font-size: 16pt; margin-bottom: 6pt; }
h2 { font-size: 13pt; margin-bottom: 4pt; }
h3 { font-size: 12pt; margin-bottom: 4pt; }
p  { margin: 0 0 6pt; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; }
td, th { border: 1px solid #bbb; padding: 4pt 6pt; font-size: 10pt; }
th { background: #f0f0f0; font-weight: bold; }
img { max-width: 100%; }
"""


def convert(input_path: str, output_path: str) -> None:
    """Convert DOCX to PDF via mammoth→HTML→xhtml2pdf. Raises on failure."""
    with open(input_path, "rb") as f:
        result = mammoth.convert_to_html(f)

    html = (
        "<!DOCTYPE html><html><head>"
        "<meta charset='utf-8'>"
        f"<style>{_CSS}</style>"
        "</head>"
        f"<body>{result.value}</body></html>"
    )

    with open(output_path, "wb") as out_f:
        status = pisa.CreatePDF(BytesIO(html.encode("utf-8")), dest=out_f)

    if status.err:
        raise RuntimeError(f"xhtml2pdf error: {status.err}")
