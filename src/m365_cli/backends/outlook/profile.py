from __future__ import annotations

import os
import winreg
from collections.abc import Iterator
from pathlib import Path
from typing import Literal

from m365_cli.core.markup import MailStyle, Signature, body_fragment, decode_html_file, parse_style

_OFFICE = r"Software\Microsoft\Office\16.0"
_MAIL_SETTINGS = _OFFICE + r"\Common\MailSettings"
_PROFILES = _OFFICE + r"\Outlook\Profiles"
_FONT_VALUES = {"compose": "ComposeFontComplex", "reply": "ReplyFontComplex"}
_SIGNATURE_VALUES = {"new": "New Signature", "reply": "Reply-Forward Signature"}

SignatureKind = Literal["new", "reply"]


def signatures_dir() -> Path:
    """Folder where Outlook keeps the signatures (%APPDATA%/Microsoft/Signatures)."""
    return Path(os.environ.get("APPDATA", str(Path.home() / "AppData/Roaming"))) / "Microsoft" / "Signatures"


def _value(key: winreg.HKEYType, name: str) -> object | None:
    try:
        return winreg.QueryValueEx(key, name)[0]
    except OSError:
        return None


def _as_text(value: object) -> str:
    if isinstance(value, bytes):
        return value.decode("utf-16-le", errors="ignore").rstrip("\x00")
    return str(value or "")


def _subkeys(root: winreg.HKEYType) -> Iterator[winreg.HKEYType]:
    index = 0
    while True:
        try:
            name = winreg.EnumKey(root, index)
        except OSError:
            return
        index += 1
        with winreg.OpenKey(root, name) as key:
            yield key
            yield from _subkeys(key)


def mail_style(kind: Literal["compose", "reply"]) -> MailStyle:
    """The user's font for new mails ('compose') or replies ('reply'); empty when not set."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _MAIL_SETTINGS) as key:
            raw = _value(key, _FONT_VALUES[kind])
    except OSError:
        return MailStyle()
    return parse_style(raw.decode("latin-1") if isinstance(raw, bytes) else str(raw or ""))


def default_signature_name(kind: SignatureKind, account: str) -> str | None:
    """Name of the signature Outlook uses for new mails or replies of `account` (None when unset)."""
    try:
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _PROFILES)
    except OSError:
        return None
    with root:
        for key in _subkeys(root):
            if _as_text(_value(key, "Account Name")).lower() == account.lower():
                name = _as_text(_value(key, _SIGNATURE_VALUES[kind]))
                return name or None
    return None


def load_signature(name: str) -> Signature | None:
    """Read signature `name` (its .htm and .txt files); None when it does not exist."""
    folder = signatures_dir()
    htm, txt = folder / f"{name}.htm", folder / f"{name}.txt"
    if not htm.exists() and not txt.exists():
        return None
    fragment = body_fragment(decode_html_file(htm.read_bytes())) if htm.exists() else ""
    text = txt.read_bytes().decode("utf-8-sig", errors="replace") if txt.exists() else ""
    if text and "\x00" in text:
        text = txt.read_bytes().decode("utf-16", errors="replace")
    return Signature(name=name, html=fragment, text=text.strip())


def signature_for(kind: SignatureKind, account: str, setting: str) -> Signature | None:
    """Signature to use: `setting` is 'auto' (Outlook's choice), 'none', or a signature name."""
    if setting.lower() == "none":
        return None
    name = default_signature_name(kind, account) if setting.lower() == "auto" else setting
    return load_signature(name) if name else None
