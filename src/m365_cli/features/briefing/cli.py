from __future__ import annotations

import typer

from m365_cli.core.feature import ContextProvider
from m365_cli.features.briefing.compose import briefing


def build_cli(get_ctx: ContextProvider) -> typer.Typer:
    """`m365 briefing` command."""
    app = typer.Typer(help="A daily 'what needs me' summary.", no_args_is_help=False)

    @app.callback(invoke_without_command=True)
    def run() -> None:
        """Show today's meetings, unread/awaiting mail, and Teams mentions."""
        typer.echo(briefing(get_ctx()))

    return app
