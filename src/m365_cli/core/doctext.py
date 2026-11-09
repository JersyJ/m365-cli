from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from pypdf import PdfReader
from pypdf.errors import PdfReadError

from m365_cli.core.errors import InvalidInputError

TEXT_SUFFIXES = frozenset(
    {".txt", ".csv", ".md", ".json", ".xml", ".log", ".yaml", ".yml", ".html", ".htm", ".ics"}
)
SUPPORTED = "PDF, DOCX, PPTX, XLSX and plain-text files"

_WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_DRAWING_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_SHEET_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_SLIDE_NUMBER = re.compile(r"ppt/slides/slide(\d+)\.xml$")

logging.getLogger("pypdf").setLevel(logging.ERROR)


def extract_text(path: Path) -> str:
    """Plain text of a PDF, DOCX, PPTX, XLSX or text file.

    Raises:
        InvalidInputError: The file type is not supported or the file cannot be read.
    """
    suffix = path.suffix.lower()
    try:
        if suffix == ".pdf":
            return "\n\n".join((page.extract_text() or "").strip() for page in PdfReader(path).pages).strip()
        if suffix == ".docx":
            return _docx(path)
        if suffix == ".pptx":
            return _pptx(path)
        if suffix == ".xlsx":
            return _xlsx(path)
        if suffix in TEXT_SUFFIXES:
            return path.read_bytes().decode("utf-8", errors="replace")
    except (PdfReadError, zipfile.BadZipFile, ElementTree.ParseError, KeyError, OSError) as exc:
        raise InvalidInputError(f"Cannot read {path.name}: {exc}") from exc
    raise InvalidInputError(f"Cannot read text from {path.name}; supported are {SUPPORTED}.")


def _xml(archive: zipfile.ZipFile, member: str) -> ElementTree.Element:
    return ElementTree.fromstring(archive.read(member))


def _docx(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        body = _xml(archive, "word/document.xml")
    paragraphs = ("".join(t.text or "" for t in p.iter(f"{_WORD_NS}t")) for p in body.iter(f"{_WORD_NS}p"))
    return "\n".join(p for p in paragraphs if p.strip())


def _pptx(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        slides = sorted(
            (int(m.group(1)), name) for name in archive.namelist() if (m := _SLIDE_NUMBER.match(name))
        )
        texts = []
        for number, name in slides:
            words = [t.text or "" for t in _xml(archive, name).iter(f"{_DRAWING_NS}t")]
            texts.append(f"[Slide {number}] " + " ".join(w for w in words if w.strip()))
    return "\n".join(texts)


def _xlsx(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            for item in _xml(archive, "xl/sharedStrings.xml").iter(f"{_SHEET_NS}si"):
                shared.append("".join(t.text or "" for t in item.iter(f"{_SHEET_NS}t")))
        lines = []
        for sheet in sorted(n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")):
            lines.append(f"[{Path(sheet).stem}]")
            for row in _xml(archive, sheet).iter(f"{_SHEET_NS}row"):
                cells = []
                for cell in row.iter(f"{_SHEET_NS}c"):
                    value = cell.find(f"{_SHEET_NS}v")
                    text = (
                        value.text
                        if value is not None
                        else "".join(t.text or "" for t in cell.iter(f"{_SHEET_NS}t"))
                    )
                    if cell.get("t") == "s" and text is not None and text.isdigit():
                        text = shared[int(text)]
                    cells.append(text or "")
                if any(cells):
                    lines.append("\t".join(cells))
    return "\n".join(lines)
