# PDF to DOCX Converter

A command-line Python utility that converts PDF files into `.docx` documents with support for interactive selection, batch processing, and three conversion modes.

## Conversion Modes

- **`layout`** (`.docx`) — Preserves visual layout using `pdf2docx` with lattice and stream table detection disabled.
- **`clean`** (`_clean.docx`) — Extracts text blocks in reading order into plain editable paragraphs and skips image blocks.
- **`smart`** (`_smart.docx`) — Uses recursive XY-cut reading order for multi-column layouts (such as CVs and brochures), preserving headings, bold formatting, and bullet items.

## Requirements & Installation

Install the required packages before running the script:

```bash
pip install PyMuPDF python-docx pdf2docx
```

## Usage

```bash
# Interactive mode (select PDF and conversion mode from prompts)
python pdftodocx_2.py

# Convert a specific PDF and choose the mode interactively
python pdftodocx_2.py report.pdf

# Convert a specific PDF in a chosen mode without prompts
python pdftodocx_2.py report.pdf -m clean

# Convert all PDFs in the current directory
python pdftodocx_2.py --all -m layout

# Scan a specific directory and show verbose debug output
python pdftodocx_2.py -d /path/to/pdfs -m smart -v
```

## Notes

- If a target `.docx` file is currently open in Microsoft Word, conversion returns a permission error until the file is closed.
