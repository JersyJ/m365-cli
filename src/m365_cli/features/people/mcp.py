from __future__ import annotations

from typing import Annotated

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from m365_cli.core.feature import ContextProvider
from m365_cli.core.mcp_support import READ_ONLY, tool_errors
from m365_cli.features.people import render
from m365_cli.features.people.api import people_api


def register_mcp(server: MCPServer, get_ctx: ContextProvider) -> None:
    """Register the people tool."""

    @server.tool(title="Look up person", annotations=READ_ONLY, structured_output=False)
    @tool_errors
    def people_lookup(
        query: Annotated[str, Field(description="Name, part of a name, or email of one person.")],
        availability_days: Annotated[
            int,
            Field(
                ge=0,
                le=14,
                description="0 = contact card only; >0 = also OOO periods and busy times for that many days.",
            ),
        ] = 0,
    ) -> str:
        """Resolve a colleague (name or email): email, title, department, phone, manager.

        With availability_days > 0 also: out-of-office periods and busy times today from the calendar
        (for scheduling). For live status (available / busy / in a meeting / away / out of office) use
        teams_status, which is more current. An ambiguous name returns an error; use the full name or email.
        """
        return render.person(*people_api(get_ctx()).lookup(query, availability_days=availability_days))
