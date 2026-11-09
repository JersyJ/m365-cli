from __future__ import annotations

import html
import re
from dataclasses import dataclass
from typing import Literal

from markdown_it import MarkdownIt

from m365_cli.core.text import plain_to_html

BodyFormat = Literal["markdown", "html", "text"]

_MARKDOWN = MarkdownIt("commonmark", {"breaks": True, "html": False}).enable("table").enable("strikethrough")
_BODY = re.compile(r"<body[^>]*>(.*)</body>", re.I | re.S)
_TAGS = re.compile(r"<[^>]+>")
_BLOCK_END = re.compile(r"</(p|div|li|tr|h[1-6])>|<br\s*/?>", re.I)
_FONT_FAMILY = re.compile(r"font-family:\s*([^;}]+)", re.I)
_FONT_SIZE = re.compile(r"(?:mso-ansi-)?font-size:\s*([\d.]+)pt", re.I)
_COLOR = re.compile(r"(?<![-\w])color:\s*(#[0-9a-f]{3,6}|rgb\([^)]*\))", re.I)
_CHARSET = re.compile(rb"charset=[\"']?([\w-]+)", re.I)


@dataclass(frozen=True, slots=True)
class MailStyle:
    """Font settings for the text we write (taken from Outlook's compose/reply font)."""

    font_family: str = ""
    font_size_pt: float | None = None
    color: str = ""

    def css(self) -> str:
        """Inline CSS for a wrapping element ('' when nothing is set)."""
        parts = []
        if self.font_family:
            parts.append(f"font-family:{self.font_family}")
        if self.font_size_pt:
            parts.append(f"font-size:{self.font_size_pt:g}pt")
        if self.color:
            parts.append(f"color:{self.color}")
        return ";".join(parts)


@dataclass(frozen=True, slots=True)
class Signature:
    """An Outlook signature: its name plus the HTML body fragment and the plain-text version."""

    name: str
    html: str
    text: str


def parse_style(css_text: str) -> MailStyle:
    """Font family, size and color from an Outlook style definition (e.g. the compose-font setting)."""
    family = _FONT_FAMILY.search(css_text)
    size = _FONT_SIZE.search(css_text)
    color = _COLOR.search(css_text)
    return MailStyle(
        font_family=html.unescape(family.group(1).strip()) if family else "",
        font_size_pt=float(size.group(1)) if size else None,
        color=color.group(1) if color else "",
    )


def to_html(text: str, body_format: BodyFormat) -> str:
    """HTML fragment for a body written as Markdown, HTML or plain text."""
    if body_format == "html":
        return text
    if body_format == "text":
        return plain_to_html(text)
    return _MARKDOWN.render(text).strip()


def to_text(text: str, body_format: BodyFormat) -> str:
    """Plain-text version of a body (Markdown and text stay as written; HTML loses its tags)."""
    if body_format != "html":
        return text
    return html.unescape(_TAGS.sub("", _BLOCK_END.sub("\n", text))).strip()


def styled(fragment: str, style: MailStyle) -> str:
    """Wrap a fragment in a div that carries the user's font."""
    css = style.css()
    return (
        f'<div style="{html.escape(css, quote=True)}">{fragment}</div>' if css else f"<div>{fragment}</div>"
    )


def html_document(fragment: str) -> str:
    """A complete HTML mail body around a fragment."""
    return f'<html><head><meta charset="utf-8"></head><body>{fragment}</body></html>'


def body_fragment(document: str) -> str:
    """The inner part of an HTML document's body (the whole text when there is no body element)."""
    m = _BODY.search(document)
    return m.group(1).strip() if m else document.strip()


def decode_html_file(raw: bytes) -> str:
    """Decode an HTML file using its declared charset (Outlook signatures are often windows-1252)."""
    declared = _CHARSET.search(raw[:2048])
    for encoding in (declared.group(1).decode() if declared else None, "utf-8", "windows-1252"):
        if encoding:
            try:
                return raw.decode(encoding)
            except (LookupError, UnicodeDecodeError):
                continue
    return raw.decode("utf-8", errors="replace")
