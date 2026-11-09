from __future__ import annotations

from typing import Annotated

import typer

from m365_cli.core.feature import ContextProvider
from m365_cli.features.people import render
from m365_cli.features.people.api import PeopleApi, people_api

Query = Annotated[str, typer.Argument(help="Name or email (resolved like in Outlook's To: field).")]


def build_cli(get_ctx: ContextProvider) -> typer.Typer:
    """`m365 people …` commands."""
    app = typer.Typer(help="Look up colleagues and their availability / out of office.", no_args_is_help=True)

    def api() -> PeopleApi:
        return people_api(get_ctx())

    @app.command()
    def find(query: Query) -> None:
        """Contact card: email, title, department, phone, manager."""
        typer.echo(render.person(*api().lookup(query)))

    @app.command()
    def status(
        query: Query, days: Annotated[int, typer.Option("--days", "-d", help="Look-ahead.")] = 7
    ) -> None:
        """Current availability, out-of-office periods and today's busy times."""
        typer.echo(render.person(*api().lookup(query, availability_days=days)))

    return app
