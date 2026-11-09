from __future__ import annotations

import io
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Annotated, Any

import typer

from m365_cli import __version__
from m365_cli.core.config import Settings
from m365_cli.core.context import AppContext
from m365_cli.core.errors import M365Error
from m365_cli.core.feature import discover_features


@dataclass
class _Run:
    """Global options and context of one CLI run."""

    overrides: dict[str, Any] = field(default_factory=dict)
    context: AppContext | None = None


_RUN = _Run()


def _context() -> AppContext:
    if _RUN.context is None:
        _RUN.context = AppContext(Settings.load(**_RUN.overrides))
    return _RUN.context


def _dev_app() -> typer.Typer:
    dev = typer.Typer(help="Developer tools: snapshot real data, reset the fake backend.")

    @dev.command()
    def snapshot(
        out: Annotated[
            Path | None, typer.Option("--out", "-o", help="Target file (default: data dir).")
        ] = None,
        mail_days: Annotated[int, typer.Option("--mail-days", help="Copy mails of the last N days.")] = 14,
        max_inbox: Annotated[int, typer.Option("--max-inbox", help="Newest Inbox mails to copy.")] = 100,
        max_people: Annotated[int, typer.Option("--max-people")] = 40,
        injection: Annotated[
            bool, typer.Option("--with-injection-mail", help="Add one synthetic prompt-injection mail.")
        ] = False,
    ) -> None:
        """Copy real data from classic Outlook into a fake-state file for evaluations.

        The file contains private data: keep it local and out of version control.
        """
        from m365_cli.backends.fake import write_state
        from m365_cli.backends.fake.snapshot import SnapshotOptions, take_snapshot

        ctx = AppContext(Settings.load(**{**_RUN.overrides, "backend": "outlook", "fake_state": None}))
        try:
            options = SnapshotOptions(
                mail_days=mail_days, max_inbox=max_inbox, max_people=max_people, add_injection_mail=injection
            )
            state = take_snapshot(ctx.outlook, options)
        finally:
            ctx.close()
        target = out or ctx.settings.data_dir / "fake-snapshot.json"
        write_state(target, state)
        typer.echo(
            f"Snapshot: {len(state.mails)} mails, {len(state.events)} events, "
            f"{len(state.people)} people -> {target}"
        )

    @dev.command("reset-fake")
    def reset_fake(path: Annotated[Path, typer.Argument(help="State file to write.")]) -> None:
        """Write the synthetic test data set to a fake-state file."""
        from m365_cli.backends.fake import write_state
        from m365_cli.backends.fake.fixtures import build_state

        write_state(path, build_state())
        typer.echo(f"Synthetic fake state -> {path}")

    return dev


def build_app() -> typer.Typer:
    """Root CLI: one sub-command per enabled feature, plus the shared commands."""
    app = typer.Typer(
        help="Microsoft 365 for agents: compact, scriptable access to your own mail and calendar.",
        no_args_is_help=True,
        pretty_exceptions_enable=False,
        rich_markup_mode=None,
    )

    @app.callback()
    def main_options(
        fake_state: Annotated[
            Path | None, typer.Option("--fake-state", help="Use the fake backend with this state file.")
        ] = None,
        data_dir: Annotated[
            Path | None, typer.Option("--data-dir", help="Directory for refs and the send log.")
        ] = None,
    ) -> None:
        """Global options."""
        _RUN.overrides = {"fake_state": fake_state, "data_dir": data_dir}

    settings = Settings.load()
    for feature in discover_features(settings):
        if feature.build_cli:
            app.add_typer(feature.build_cli(_context), name=feature.name, help=feature.help)
    app.add_typer(_dev_app(), name="dev")

    @app.command()
    def mcp() -> None:
        """Run the MCP server on stdio (used by Claude Code / other MCP clients)."""
        from m365_cli.mcp_server import run

        run(_context())

    @app.command()
    def log(limit: Annotated[int, typer.Option("--limit", "-n")] = 20) -> None:
        """Show the latest send actions (mail, invitations, responses, cancellations)."""
        entries = _context().audit.recent(limit)
        typer.echo("\n".join(f"{e.at} {e.action:12} {e.summary}" for e in entries) or "(no send actions)")

    @app.command()
    def features() -> None:
        """List enabled features."""
        for f in discover_features(_context().settings):
            typer.echo(f"{f.name:10} {f.help}")

    @app.command()
    def version() -> None:
        """Print the version."""
        typer.echo(__version__)

    return app


def main() -> None:
    """Entry point: run the CLI; expected errors are printed as one line with exit code 1."""
    for stream in (sys.stdout, sys.stderr):
        if isinstance(stream, io.TextIOWrapper):
            stream.reconfigure(encoding="utf-8")
    try:
        build_app()()
    except M365Error as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise SystemExit(1) from None
    finally:
        if _RUN.context is not None:
            _RUN.context.close()
