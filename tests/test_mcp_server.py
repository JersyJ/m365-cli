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

SEND_TOOLS = {"mail_send", "cal_send", "cal_respond", "cal_cancel", "teams_send"}
TOOL_COUNT = 24


def call(ctx: AppContext, name: str, **arguments: Any) -> str:
    result = asyncio.run(build_server(ctx).call_tool(name, arguments))
    assert isinstance(result, CallToolResult)
    return "\n".join(c.text for c in result.content if isinstance(c, TextContent))


def test_features_are_discovered_and_can_be_disabled() -> None:
    assert [f.name for f in discover_features(Settings())] == ["calendar", "mail", "people", "teams"]
    assert [
        f.name for f in discover_features(Settings(disabled_features=["people", "calendar", "teams"]))
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
