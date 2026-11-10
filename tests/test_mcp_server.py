import asyncio
from pathlib import Path
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult, TextContent

from m365_cli.core.config import Settings
from m365_cli.core.context import AppContext
from m365_cli.core.feature import discover_features
from m365_cli.mcp_server import build_server

SEND_TOOLS = {"mail_send", "cal_send", "cal_respond", "cal_cancel", "teams_send", "teams_update"}
TOOL_COUNT = 31


def call(ctx: AppContext, name: str, **arguments: Any) -> str:
    result = asyncio.run(build_server(ctx).call_tool(name, arguments))
    assert isinstance(result, CallToolResult)
    return "\n".join(c.text for c in result.content if isinstance(c, TextContent))


def test_features_are_discovered_and_can_be_disabled() -> None:
    assert [f.name for f in discover_features(Settings())] == [
        "briefing",
        "calendar",
        "mail",
        "people",
        "teams",
    ]
    assert [
        f.name
        for f in discover_features(Settings(disabled_features=["people", "calendar", "teams", "briefing"]))
    ] == ["mail"]


def test_tool_metadata(fake_ctx: AppContext) -> None:
    tools = {t.name: t for t in asyncio.run(build_server(fake_ctx).list_tools())}
    assert len(tools) == TOOL_COUNT
    for name, tool in tools.items():
        assert tool.title, name
        assert tool.annotations is not None, name
        forced = (tool.meta or {}).get("anthropic/requiresUserInteraction")
        assert bool(forced) == (name in SEND_TOOLS), name
    assert tools["mail_find"].input_schema["properties"]["sender"]["description"]


def test_mcp_server_does_not_start_outlook(tmp_path: Path) -> None:
    ctx = AppContext(Settings(data_dir=tmp_path))
    asyncio.run(build_server(ctx).list_tools())
    assert "outlook" not in ctx.__dict__
    assert "teams" not in ctx.__dict__


def test_find_read_and_untrusted_body(fake_ctx: AppContext) -> None:
    listing = call(fake_ctx, "mail_find", sender="billing")
    ref = listing.split()[0]
    body = call(fake_ctx, "mail_read", mail_ref=ref)
    assert "<<<UNTRUSTED MAIL BODY" in body and "payments@evil.example" in body


def test_draft_then_send_and_paging_hint(fake_ctx: AppContext) -> None:
    ref = call(fake_ctx, "mail_find", subject="runner capacity").split()[0]
    preview = call(fake_ctx, "mail_draft", mode="reply", mail_ref=ref, body="Friday.")
    assert "NOT SENT" in preview and "mail_send(draft_ref=" in preview
    draft_ref = preview.split()[1]
    assert call(fake_ctx, "mail_send", draft_ref=draft_ref).startswith("Sent")
    assert "offset=1" in call(fake_ctx, "mail_find", limit=1)


def test_errors_are_tool_errors(fake_ctx: AppContext) -> None:
    with pytest.raises(ToolError, match="Unknown ref"):
        call(fake_ctx, "mail_read", mail_ref="m:NOPE123")
    with pytest.raises(ToolError, match="Unrecognized date"):
        call(fake_ctx, "cal_find", start="someday")


def test_calendar_and_people_tools(fake_ctx: AppContext) -> None:
    assert "08:00-09:00" in call(
        fake_ctx, "cal_find_free_time", attendees=["Jana", "Petr"], date="tomorrow", duration_minutes=60
    )
    assert "Out of office" in call(fake_ctx, "people_lookup", query="Eva", availability_days=5)


def _handshake_response(frames: list[dict[str, Any]], init_id: int) -> dict[str, Any]:
    """Drive the stdio server in-process with the given frames; return the init response."""
    import json

    import anyio
    from mcp import types as mcp_types_mod
    from mcp.server.runner import serve_loop
    from mcp.shared.message import SessionMessage

    from m365_cli.mcp_server import build_server

    ctx = AppContext(Settings())
    lowlevel = build_server(ctx)._lowlevel_server
    init_options = lowlevel.create_initialization_options()
    result: dict[str, Any] = {}

    async def main() -> None:
        to_server_send, to_server_recv = anyio.create_memory_object_stream[SessionMessage | Exception](8)
        from_server_send, from_server_recv = anyio.create_memory_object_stream[SessionMessage](8)
        for frame in frames:
            message = mcp_types_mod.jsonrpc_message_adapter.validate_python(frame, by_name=False)
            await to_server_send.send(SessionMessage(message))
        await to_server_send.aclose()
        async with lowlevel.lifespan(lowlevel) as state:
            with anyio.move_on_after(5):
                async with anyio.create_task_group() as tg:

                    async def reader() -> None:
                        async for message in from_server_recv:
                            payload = json.loads(message.message.model_dump_json(by_alias=True))
                            if payload.get("id") == init_id:
                                result.update(payload)
                                tg.cancel_scope.cancel()

                    tg.start_soon(reader)
                    await serve_loop(
                        lowlevel,
                        to_server_recv,
                        from_server_send,
                        lifespan_state=state,
                        init_options=init_options,
                    )

    try:
        anyio.run(main)
    finally:
        ctx.close()
    return result


def test_handshake_survives_a_modern_probe_frame() -> None:
    """A modern-envelope probe before `initialize` must not break the handshake (regression)."""
    modern_probe = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "ping",
        "params": {"_meta": {"io.modelcontextprotocol/protocolVersion": "2026-07-28"}},
    }
    initialize = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "initialize",
        "params": {
            "protocolVersion": "2025-11-25",
            "capabilities": {},
            "clientInfo": {"name": "t", "version": "1"},
        },
    }
    response = _handshake_response([modern_probe, initialize], init_id=2)
    assert "result" in response, response
