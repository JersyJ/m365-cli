from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from m365_cli.core.feature import ContextProvider
from m365_cli.core.mcp_support import READ_ONLY, tool_errors
from m365_cli.features.briefing.compose import briefing as compose_briefing


def register_mcp(server: MCPServer, get_ctx: ContextProvider) -> None:
    """Register the briefing tool."""

    @server.tool(title="Daily briefing", annotations=READ_ONLY, structured_output=False)
    @tool_errors
    def briefing() -> str:
        """One 'what needs me today' summary: today's meetings, unread/awaiting mail, and Teams mentions.

        Use it to start the day or answer "what do I need to deal with". Combines calendar, mail and Teams.
        """
        return compose_briefing(get_ctx())
