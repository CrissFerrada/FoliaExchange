import pytest
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from docx import Document


@pytest.fixture
def sample_pdf(tmp_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)
    pdf.cell(200, 10, text="Hello FoliaExchange test.",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    path = tmp_path / "sample.pdf"
    pdf.output(str(path))
    return str(path)


@pytest.fixture
def sample_docx(tmp_path):
    doc = Document()
    doc.add_heading("FoliaExchange Test", 0)
    doc.add_paragraph("Hello, this is a test document.")
    path = tmp_path / "sample.docx"
    doc.save(str(path))
    return str(path)
