#!/usr/bin/env python3
"""PDF to DOCX converter.

Converts PDF files to Word documents using one of two modes:

* layout - preserves visual layout (via pdf2docx), with table detection disabled.
* clean  - extracts plain text only, producing simple editable paragraphs.
* smart  - column-aware reading order with headings/bold (best for CVs).

Usage examples:
    python index.py                       # interactive mode
    python index.py report.pdf            # prompts for mode
    python index.py report.pdf -m clean   # fully non-interactive
    python index.py --all -m layout       # convert every PDF in the folder
"""

from __future__ import annotations

import argparse
import logging
import sys
import warnings
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import List, Optional, Sequence

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt
from pdf2docx import Converter

# Silence noisy deprecation warnings coming from third-party libraries.
warnings.filterwarnings("ignore", category=FutureWarning)

LOGGER = logging.getLogger("pdf2docx_cli")
TEXT_BLOCK_TYPE = 0  # PyMuPDF block type for text (1 = image)


class ConversionMode(Enum):
    """Supported conversion strategies."""

    LAYOUT = "layout"
    CLEAN = "clean"
    SMART = "smart"

    @property
    def suffix(self) -> str:
        return {
            ConversionMode.LAYOUT: ".docx",
            ConversionMode.CLEAN: "_clean.docx",
            ConversionMode.SMART: "_smart.docx",
        }[self]


@dataclass(frozen=True)
class ConversionResult:
    """Outcome of a single file conversion."""

    source: Path
    output: Optional[Path]
    error: Optional[str] = None

    @property
    def ok(self) -> bool:
        return self.error is None


class PDFToDocxConverter:
    """Discovers PDF files and converts them to DOCX."""

    def __init__(self, directory: Path = Path(".")) -> None:
        self.directory = directory

    # ------------------------------------------------------------------ discovery
    def find_pdfs(self) -> List[Path]:
        """Return a sorted list of PDF files in the working directory."""
        return sorted(
            p for p in self.directory.iterdir()
            if p.is_file() and p.suffix.lower() == ".pdf"
        )

    # ----------------------------------------------------------------- conversion
    @staticmethod
    def output_path_for(pdf_path: Path, mode: ConversionMode) -> Path:
        return pdf_path.with_name(f"{pdf_path.stem}{mode.suffix}")

    def convert(self, pdf_path: Path, mode: ConversionMode) -> ConversionResult:
        """Convert one PDF, returning a result object instead of raising."""
        output_path = self.output_path_for(pdf_path, mode)
        try:
            if mode is ConversionMode.LAYOUT:
                self._convert_with_layout(pdf_path, output_path)
            elif mode is ConversionMode.SMART:
                self._convert_smart(pdf_path, output_path)
            else:
                self._convert_clean_text(pdf_path, output_path)
        except PermissionError:
            return ConversionResult(
                pdf_path, None,
                f"Permission denied. Close '{output_path.name}' in MS Word and retry.",
            )
        except Exception as exc:  # noqa: BLE001 - report any library failure cleanly
            LOGGER.debug("Conversion failed", exc_info=True)
            return ConversionResult(pdf_path, None, f"{type(exc).__name__}: {exc}")
        return ConversionResult(pdf_path, output_path)

    @staticmethod
    def _convert_with_layout(pdf_path: Path, output_path: Path) -> None:
        """Preserve visual layout while minimising table boxes."""
        converter = Converter(str(pdf_path))
        try:
            converter.convert(
                str(output_path),
                parse_lattice_table=False,
                parse_stream_table=False,
            )
        finally:
            converter.close()

    @staticmethod
    def _convert_clean_text(pdf_path: Path, output_path: Path) -> None:
        """Extract text blocks in reading order into plain paragraphs."""
        word_doc = Document()
        with fitz.open(str(pdf_path)) as pdf_doc:
            last_page = len(pdf_doc) - 1
            for page_number, page in enumerate(pdf_doc):
                for block in page.get_text("blocks", sort=True):
                    if block[6] != TEXT_BLOCK_TYPE:
                        continue  # skip image blocks
                    text = block[4].strip()
                    if text:
                        word_doc.add_paragraph(text)
                if page_number != last_page:
                    word_doc.add_page_break()
        word_doc.save(str(output_path))

    @staticmethod
    def _convert_smart(pdf_path: Path, output_path: Path) -> None:
        """Column-aware conversion: correct reading order for multi-column PDFs
        (CVs, brochures) with headings and bold text preserved, icons dropped."""
        word_doc = Document()
        section = word_doc.sections[0]
        section.left_margin = section.right_margin = Inches(0.8)
        section.top_margin = section.bottom_margin = Inches(0.7)

        with fitz.open(str(pdf_path)) as pdf_doc:
            last_page = len(pdf_doc) - 1
            for page_number, page in enumerate(pdf_doc):
                blocks = [
                    b for b in page.get_text("dict")["blocks"]
                    if b["type"] == TEXT_BLOCK_TYPE
                ]
                for block in xy_cut_order(blocks):
                    _write_block(word_doc, block)
                if page_number != last_page:
                    word_doc.add_page_break()
        word_doc.save(str(output_path))


# -------------------------------------------------------------- reading-order helpers
MIN_GAP_Y = 5.0   # small vertical gap (pt) that separates two bands
BAND_GAP_Y = 25.0 # large vertical gap: header / body separation
MIN_GAP_X = 10.0  # horizontal white-space (pt) that separates two columns
BOLD_FLAG = 16    # PyMuPDF span flag for bold
BULLET_CHARS = "\u2022\u00b7\u25cf\u25aa-"


def _find_gap(spans: List[tuple], min_gap: float) -> Optional[float]:
    """Return a cut position inside the first empty gap of 1-D intervals."""
    spans = sorted(spans)
    reach = spans[0][1]
    for start, end in spans[1:]:
        if start - reach >= min_gap:
            return (start + reach) / 2
        reach = max(reach, end)
    return None


def _split(blocks: list, axis: int, cut: float) -> tuple:
    """Split blocks into (before, after) a cut position along x (0) or y (1)."""
    before = [b for b in blocks if (b["bbox"][axis] + b["bbox"][axis + 2]) / 2 < cut]
    after = [b for b in blocks if b not in before]
    return before, after


def xy_cut_order(blocks: list) -> list:
    """Recursive XY-cut: order blocks the way a person reads a multi-column page.

    Order of attempts: large horizontal gap (header/body) -> column split ->
    small horizontal gap -> plain top-to-bottom sort.
    """
    if len(blocks) <= 1:
        return blocks

    y_spans = [(b["bbox"][1], b["bbox"][3]) for b in blocks]
    x_spans = [(b["bbox"][0], b["bbox"][2]) for b in blocks]

    for axis, spans, gap in ((1, y_spans, BAND_GAP_Y), (0, x_spans, MIN_GAP_X),
                             (1, y_spans, MIN_GAP_Y)):
        cut = _find_gap(spans, gap)
        if cut is not None:
            first, second = _split(blocks, axis, cut)
            return xy_cut_order(first) + xy_cut_order(second)

    return sorted(blocks, key=lambda b: (b["bbox"][1], b["bbox"][0]))


def _line_info(line: dict) -> tuple:
    """Return (text, is_bold, max_font_size) for one PDF text line."""
    spans = [s for s in line["spans"] if s["text"].strip()]
    text = "".join(s["text"] for s in line["spans"]).replace("\u200b", "").strip()
    bold = bool(spans) and all(s["flags"] & BOLD_FLAG or "bold" in s["font"].lower() for s in spans)
    size = max((s["size"] for s in spans), default=0)
    return text, bold, size


def _write_block(word_doc: Document, block: dict) -> None:
    """Write a PDF text block as headings / paragraphs, keeping bold text."""
    lines = [_line_info(line) for line in block["lines"]]
    lines = [info for info in lines if info[0]]
    if not lines:
        return

    # Short all-caps single line => section heading.
    if len(lines) == 1 and lines[0][0].isupper() and len(lines[0][0]) < 40:
        text, _, size = lines[0]
        level = 0 if size >= 20 else 2
        word_doc.add_heading(text, level=level)
        return

    # Merge wrapped lines that share the same weight; break on bold changes/bullets.
    paragraphs: List[list] = []  # [text, bold]
    for text, bold, _ in lines:
        is_bullet = text[0] in BULLET_CHARS and len(text) <= 2
        if is_bullet:
            paragraphs.append(["\u2022", False])
            continue
        if paragraphs and paragraphs[-1][1] == bold and paragraphs[-1][0] != "\u2022":
            joiner = "" if paragraphs[-1][0].endswith("-") else " "
            paragraphs[-1][0] += joiner + text
        elif paragraphs and paragraphs[-1][0] == "\u2022":
            paragraphs[-1] = [text, bold, True]  # bullet item
        else:
            paragraphs.append([text, bold])

    for item in paragraphs:
        if item[0] == "\u2022":
            continue
        style = "List Bullet" if len(item) == 3 else None
        para = word_doc.add_paragraph(style=style)
        para.paragraph_format.space_after = Pt(2)
        run = para.add_run(item[0])
        run.bold = item[1]

def prompt_choice(prompt: str, valid: Sequence[str]) -> str:
    """Ask until the user enters one of the valid answers."""
    while True:
        answer = input(prompt).strip().lower()
        if answer in valid:
            return answer
        print(f"[!] Invalid input. Choose one of: {', '.join(valid)}")


def select_files(files: List[Path]) -> List[Path]:
    """Show a numbered list and let the user pick one file or 'a' for all."""
    print("\n--- Available PDF Files ---")
    for idx, path in enumerate(files, start=1):
        print(f"  [{idx}] {path.name}")
    print("  [a] Convert all")

    valid = [str(i) for i in range(1, len(files) + 1)] + ["a"]
    choice = prompt_choice("\nSelect file number: ", valid)
    return files if choice == "a" else [files[int(choice) - 1]]


def select_mode() -> ConversionMode:
    print("\nSelect conversion mode:")
    print("  [1] Layout - preserves styling, fewer table boxes")
    print("  [2] Clean  - plain editable paragraphs, no boxes")
    print("  [3] Smart  - fixes multi-column PDFs (CVs): correct order, headings, bold")
    choice = prompt_choice("Enter mode (1, 2 or 3): ", ["1", "2", "3"])
    return {
        "1": ConversionMode.LAYOUT,
        "2": ConversionMode.CLEAN,
        "3": ConversionMode.SMART,
    }[choice]


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert PDF files to DOCX (layout-preserving or clean text)."
    )
    parser.add_argument("files", nargs="*", type=Path, help="PDF file(s) to convert")
    parser.add_argument(
        "-m", "--mode", choices=[m.value for m in ConversionMode],
        help="conversion mode (prompted if omitted)",
    )
    parser.add_argument("-a", "--all", action="store_true",
                        help="convert every PDF in the current directory")
    parser.add_argument("-d", "--directory", type=Path, default=Path("."),
                        help="directory to scan for PDFs (default: current)")
    parser.add_argument("-v", "--verbose", action="store_true", help="show debug output")
    return parser.parse_args(argv)


def resolve_targets(args: argparse.Namespace, converter: PDFToDocxConverter) -> List[Path]:
    """Work out which PDFs to convert from CLI args or interactive input."""
    if args.files:
        missing = [f for f in args.files if not f.is_file()]
        if missing:
            raise FileNotFoundError(", ".join(str(f) for f in missing))
        return list(args.files)

    pdfs = converter.find_pdfs()
    if not pdfs:
        raise FileNotFoundError(f"No PDF files found in '{converter.directory.resolve()}'.")
    return pdfs if args.all else select_files(pdfs)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )

    converter = PDFToDocxConverter(args.directory)

    try:
        targets = resolve_targets(args, converter)
        mode = ConversionMode(args.mode) if args.mode else select_mode()
    except FileNotFoundError as exc:
        LOGGER.error("[!] %s", exc)
        return 1
    except (KeyboardInterrupt, EOFError):
        LOGGER.info("\n[!] Cancelled.")
        return 130

    failures = 0
    for pdf in targets:
        LOGGER.info("[*] Converting '%s' (%s mode)...", pdf.name, mode.value)
        result = converter.convert(pdf, mode)
        if result.ok:
            LOGGER.info("[+] Saved as '%s'.", result.output.name)
        else:
            failures += 1
            LOGGER.error("[!] %s: %s", pdf.name, result.error)

    LOGGER.info("\nDone: %d succeeded, %d failed.", len(targets) - failures, failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
