# docx-from-pdf

A Python library and CLI for converting PDF files into editable DOCX documents.

`docx-from-pdf` uses PDFium for PDF parsing and rendering. It converts text,
tables, images, vector paths, and section layout into DOCX output with
permissive runtime dependencies.

This project uses PDFium through `pypdfium2`; it is not affiliated with PDFium,
Chromium, or Google.

## Install

Install from PyPI:

```bash
pip install docx-from-pdf
```

Install from source:

```bash
pip install git+https://github.com/michael-provision/docx-from-pdf.git
```

Install local development dependencies:

```bash
uv venv
uv pip install -r requirements.txt
```

## Usage

Convert a PDF from Python:

```python
from docx_from_pdf import Converter

converter = Converter("input.pdf")
converter.convert("output.docx")
converter.close()
```

Convert a PDF from the CLI:

```bash
docx-from-pdf convert input.pdf output.docx
```

Extract tables:

```python
from docx_from_pdf import Converter

converter = Converter("input.pdf")
tables = converter.extract_tables()
converter.close()
```

## Development

Run the test suite:

```bash
PYTHONPATH=. uv run --no-project --with-requirements requirements.txt --with pytest pytest -q test/test.py
```

Some visual-comparison tests render DOCX files through LibreOffice on non-Windows
systems, so LibreOffice must be installed for the full suite.

## License

MIT
