from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from m365_cli import __version__
from m365_cli.core.context import AppContext
from m365_cli.core.feature import discover_features
from m365_cli.core.mcp_support import UNTRUSTED_NOTE

BASE_INSTRUCTIONS = (
    "Access to the user's own Microsoft 365 mailbox and calendar. Results are compact text; pass refs back "
    "exactly as shown. Never send anything (mail, invitations, responses) without explicit user approval. "
    + UNTRUSTED_NOTE
)


def build_server(ctx: AppContext) -> MCPServer:
    """MCP server with the tools and instructions of all enabled features."""
    features = discover_features(ctx.settings)
    instructions = "\n".join([BASE_INSTRUCTIONS, *(f.instructions for f in features if f.instructions)])
    server = MCPServer(name="m365", version=__version__, instructions=instructions)
    for feature in features:
        if feature.register_mcp:
            feature.register_mcp(server, lambda: ctx)
    return server


def run(ctx: AppContext) -> None:
    """Run the MCP server on stdio until the client disconnects."""
    try:
        build_server(ctx).run("stdio")
    finally:
        ctx.close()
