from pdf2docx import Converter


def convert(input_path: str, output_path: str) -> None:
    """Convert PDF to DOCX using pdf2docx. Raises on failure."""
    cv = Converter(input_path)
    try:
        cv.convert(output_path, start=0, end=None)
    finally:
        cv.close()
