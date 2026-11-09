from __future__ import annotations

import sys
from collections.abc import Callable

import typer


def split_list(values: list[str] | None) -> list[str]:
    """Accept repeated options and/or comma/semicolon separated values: -t a,b -t c -> [a, b, c]."""
    out: list[str] = []
    for v in values or []:
        out += [p.strip() for p in v.replace(";", ",").split(",") if p.strip()]
    return out


def read_body(body: str) -> str:
    """'-' reads the body from stdin (handy for multi-line text)."""
    return sys.stdin.read() if body == "-" else body


def confirm_send(
    preview: str, send: Callable[[], str], *, yes: bool, keep_draft: bool, send_command: str
) -> None:
    """Print the preview, then send only after confirmation (or with --yes). --draft keeps it unsent.

    `send_command` is shown as the hint for sending the draft later.
    """
    typer.echo(preview)
    if keep_draft:
        typer.echo(f"\nKept as draft. Send later with: {send_command}")
        return
    if not yes and not typer.confirm("\nSend now?", default=False):
        typer.echo(f"Not sent; draft kept. Send later with: {send_command}")
        return
    typer.echo(send())
