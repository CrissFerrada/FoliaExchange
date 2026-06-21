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
