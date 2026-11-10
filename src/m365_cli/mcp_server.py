from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from m365_cli import __version__
from m365_cli.core.context import AppContext
from m365_cli.core.feature import discover_features
from m365_cli.core.mcp_support import UNTRUSTED_NOTE

BASE_INSTRUCTIONS = (
    "Access to the user's own Microsoft 365 mailbox, calendar and Teams. Results are compact text; pass refs "
    "back exactly as shown. Never send anything (mail, Teams messages, invitations, responses) without "
    "explicit user approval. " + UNTRUSTED_NOTE
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
        _run_stdio_handshake_only(build_server(ctx))
    finally:
        ctx.close()


def _run_stdio_handshake_only(server: MCPServer) -> None:
    """Serve stdio in handshake-only mode, so a client's modern-envelope probe cannot lock it.

    The default `run("stdio")` serves both the legacy initialize handshake and the
    2026-07-28 per-request-envelope era, and the client's first frame decides which.
    Copilot CLI sometimes sends a probe frame before `initialize`, which flips the
    connection to the modern era and then rejects the handshake (JSON-RPC -32022).
    Clients today still use the handshake, so serving it alone is both safe and stable.
    Falls back to the default if the internal mcp API changes.
    """
    import anyio

    try:
        from mcp.server.runner import serve_loop
        from mcp.server.stdio import stdio_server
    except ImportError:
        server.run("stdio")
        return

    lowlevel = server._lowlevel_server
    init_options = lowlevel.create_initialization_options()

    async def main() -> None:
        async with (
            stdio_server() as (read_stream, write_stream),
            lowlevel.lifespan(lowlevel) as lifespan_state,
        ):
            await serve_loop(
                lowlevel,
                read_stream,
                write_stream,
                lifespan_state=lifespan_state,
                init_options=init_options,
            )

    anyio.run(main)
