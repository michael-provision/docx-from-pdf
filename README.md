# PDF to DOCX Converter

A Python library and CLI for converting PDF files into editable DOCX documents.

This fork uses PDFium for PDF parsing and rendering. It keeps the original layout
pipeline for text blocks, tables, images, vector paths, sections, and DOCX
generation, but removes the PyMuPDF runtime dependency.

## Install

Install from this repository:

```bash
pip install git+https://github.com/michael-provision/pdf2docx.git
```

Install local development dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Convert a PDF from Python:

```python
from pdf2docx import Converter

converter = Converter("input.pdf")
converter.convert("output.docx")
converter.close()
```

Convert a PDF from the CLI:

```bash
pdf2docx convert input.pdf output.docx
```

Extract tables:

```python
from pdf2docx import Converter

converter = Converter("input.pdf")
tables = converter.extract_tables()
converter.close()
```

## Development

Run the test suite:

```bash
PYTHONPATH=. pytest -q test/test.py
```

Some visual-comparison tests render DOCX files through LibreOffice on non-Windows
systems, so LibreOffice must be installed for the full suite.

## License

MIT
