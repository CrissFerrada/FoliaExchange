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
