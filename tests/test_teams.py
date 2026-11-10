import asyncio
import json
from datetime import datetime, timedelta
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ToolError
from mcp_types import CallToolResult, TextContent

from m365_cli.core.config import Settings
from m365_cli.core.context import AppContext
from m365_cli.core.errors import InvalidInputError
from m365_cli.features.teams import rules
from m365_cli.features.teams.api import teams_api
from m365_cli.features.teams.models import Member
from m365_cli.mcp_server import build_server

NOW_MS = 1_789_000_000_000
ME = "8:orgid:aaaaaaaa-0000-0000-0000-000000000001"
OTHER = "8:orgid:bbbbbbbb-0000-0000-0000-000000000002"


def call(ctx: AppContext, name: str, **arguments: Any) -> str:
    result = asyncio.run(build_server(ctx).call_tool(name, arguments))
    assert isinstance(result, CallToolResult)
    return "\n".join(c.text for c in result.content if isinstance(c, TextContent))


def conv(**overrides: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "id": "19:x@thread.v2",
        "threadProperties": {"productThreadType": "Chat", "topic": "Ops"},
        "memberProperties": {"memberExpirationTime": 0},
        "properties": {"consumptionhorizon": "100;100;1"},
        "lastMessage": {"id": "200", "from": f"https://x/contacts/{OTHER}", "messagetype": "RichText/Html"},
    }
    return base | overrides


def test_system_streams_deleted_and_left_threads_are_hidden() -> None:
    assert rules.is_listed(conv(), now_ms=NOW_MS)
    assert not rules.is_listed(conv(id="48:notifications"), now_ms=NOW_MS)
    deleted = conv(threadProperties={"productThreadType": "TeamsTeam", "isdeleted": True})
    assert not rules.is_listed(deleted, now_ms=NOW_MS)
    left = conv(
        threadProperties={"productThreadType": "TeamsStandardChannel"},
        memberProperties={"memberExpirationTime": NOW_MS - 1000},
    )
    assert not rules.is_listed(left, now_ms=NOW_MS)
    expired_meeting = conv(
        threadProperties={"productThreadType": "Meeting"},
        memberProperties={"memberExpirationTime": NOW_MS - 1000},
    )
    assert rules.is_listed(expired_meeting, now_ms=NOW_MS)
    forever = conv(memberProperties={"memberExpirationTime": rules.NO_EXPIRY_MS})
    assert rules.is_listed(forever, now_ms=NOW_MS)


def test_unread_and_one_on_one_partner() -> None:
    me = Member(name="Alex", mri=ME)
    assert rules.is_unread(conv(), me)
    assert not rules.is_unread(conv(properties={"consumptionhorizon": "300;300;1"}), me)
    mine = conv(lastMessage={"id": "200", "from": f"https://x/contacts/{ME}", "messagetype": "RichText/Html"})
    assert not rules.is_unread(mine, me)
    by_name = conv(
        lastMessage={"id": "200", "from": "x/8:orgid:old", "imdisplayname": "alex", "messagetype": "Text"}
    )
    assert not rules.is_unread(by_name, me)
    one = f"19:{ME.rsplit(':', 1)[1]}_{OTHER.rsplit(':', 1)[1]}@unq.gbl.spaces"
    assert rules.other_member_mri(one, ME) == OTHER
    assert rules.other_member_mri(one, "8:orgid:cccccccc-0000-0000-0000-000000000003") == ""
    assert rules.chat_kind({"id": one, "threadProperties": {}}) == "chat"


def test_reactions_accept_string_and_list_and_skip_follow() -> None:
    as_list = [{"key": "like", "users": [{"mri": OTHER}]}, {"key": "follow", "users": [{"mri": ME}]}]
    assert rules.reactions_of({"emotions": as_list}) == {"like": 1}
    assert rules.reactions_of({"emotions": json.dumps(as_list)}) == {"like": 1}
    assert rules.reactions_of({"emotions": [{"key": "heart", "users": []}]}) == {}
    assert rules.reactions_of({"emotions": "not json"}) == {}


def test_html_to_text_keeps_mentions_quote_links_and_lists() -> None:
    content = (
        rules.quote_html("123", OTHER, "Jana", "old text")
        + rules.mention_html("Petr", 0)
        + ' see <a href="https://example.com/x">the plan</a><ul><li>one</li><li>two</li></ul>'
        + '<img src="https://x/img"><emoji alt="🙂"></emoji>'
    )
    text, mentions, quote, images = rules.html_to_text(content)
    assert mentions == ["Petr"]
    assert quote == "Jana: old text"
    assert "@Petr see the plan (https://example.com/x)" in text
    assert "- one\n- two" in text
    assert images == 1 and "🙂" in text


def test_body_html_mentions_and_no_raw_html() -> None:
    jana = Member(name="Jana Nováková", mri=OTHER)
    html, text = rules.body_html("Hi @[Jana] <b>bold?</b> **yes**", "markdown", [jana])
    assert 'itemtype="http://schema.skype.com/Mention"' in html and "Jana Nováková" in html
    assert "<b>" not in html and "&lt;b&gt;" in html and "<strong>yes</strong>" in html
    assert text.startswith("Hi @Jana")
    assert rules.resolve_mentions("Hi @[jana]", [jana]) == [jana]
    with pytest.raises(InvalidInputError):
        rules.resolve_mentions("Hi @[Nobody]", [jana])


def test_payload_has_display_name_and_mentions() -> None:
    jana = Member(name="Jana", mri=OTHER)
    body = rules.message_payload("<p>x</p>", display_name="Alex", mentions=[jana], client_id="1")
    assert body["imdisplayname"] == "Alex" and body["messagetype"] == "RichText/Html"
    assert json.loads(body["properties"]["mentions"])[0]["mri"] == OTHER
    scheduled = rules.scheduled_payload("19:x", body, datetime(2026, 11, 10, 9, 0))
    assert scheduled["draftType"] == "ScheduledDraft" and scheduled["message"]["conversationid"] == "19:x"


def test_protected_chats() -> None:
    rules.check_protected("Platform squad", ["All Company*"])
    with pytest.raises(InvalidInputError):
        rules.check_protected("All Company news", ["all company*"])


def test_fake_find_hides_deleted_and_marks_unread(fake_ctx: AppContext) -> None:
    listing = call(fake_ctx, "teams_find")
    assert "Old Project" not in listing
    assert "Jana Nováková" in listing and "*unread*" in listing
    assert "Platform squad" in call(fake_ctx, "teams_find", kind="group")


def test_fake_read_wraps_untrusted_and_pages(fake_ctx: AppContext) -> None:
    page = call(fake_ctx, "teams_read", chat="Platform squad")
    assert "<<<UNTRUSTED teams messages" in page and "SYSTEM NOTE" in page
    assert "files: deploy-key.txt" in page
    short = call(fake_ctx, "teams_read", chat="Jana", limit=2)
    assert 'older="o:' in short
    older = short.split('older="')[1].split('"')[0]
    rest = call(fake_ctx, "teams_read", chat="Jana", limit=2, older=older)
    assert "Terraform plan" in rest and "Pipeline run" not in rest


def test_fake_draft_reply_then_send(fake_ctx: AppContext) -> None:
    page = call(fake_ctx, "teams_read", chat="Jana")
    lines = page.splitlines()
    ref = next(lines[i - 1].split("]")[0].lstrip("[") for i, line in enumerate(lines) if "NAT" in line)
    preview = call(fake_ctx, "teams_draft", reply_to=ref, body="Yes, approve it. @[Jana Nováková]")
    assert "NOT SENT" in preview and "Mentions: Jana Nováková" in preview and "Reply to:" in preview
    assert not fake_ctx.fake.state.outbox
    draft_ref = preview.split()[2]
    assert call(fake_ctx, "teams_send", draft_ref=draft_ref).startswith("Sent to Jana Nováková")
    out = fake_ctx.fake.state.outbox[-1]
    assert out.action == "teams_send" and "approve it" in out.summary
    with pytest.raises(ToolError, match="Unknown Teams draft"):
        asyncio.run(build_server(fake_ctx).call_tool("teams_send", {"draft_ref": draft_ref}))


def test_fake_new_chat_by_email_and_schedule(fake_ctx: AppContext) -> None:
    api = teams_api(fake_ctx)
    draft = api.draft("Hello", to="petr.svoboda@example.com")
    assert draft.chat_id == "" and draft.to is not None
    when = fake_ctx.fake.state.clock + timedelta(days=1)
    assert "Scheduled" in api.send(draft.ref, send_at=when)
    assert fake_ctx.fake.state.outbox[-1].send_at == when


def test_fake_protected_chat_blocks_send(tmp_path: Any) -> None:
    settings = Settings(
        backend="fake", fake_state=tmp_path / "s.json", data_dir=tmp_path, teams_protected=["Platform*"]
    )
    ctx = AppContext(settings)
    try:
        api = teams_api(ctx)
        draft = api.draft("hi", chat="Platform squad")
        with pytest.raises(InvalidInputError):
            api.send(draft.ref)
        assert not ctx.fake.state.outbox
    finally:
        ctx.close()


def test_fake_status_meeting_and_ooo(fake_ctx: AppContext) -> None:
    jana = call(fake_ctx, "teams_status", who="Jana")
    assert "Status: In a meeting" in jana
    assert "Heads-down" in jana and "Cloud Engineer" in jana
    eva = call(fake_ctx, "teams_status", who="eva.horakova@example.com")
    assert "Out of office" in eva and "On leave" in eva


def test_fake_status_ambiguous_and_unknown(fake_ctx: AppContext) -> None:
    with pytest.raises(ToolError, match="No Teams user"):
        asyncio.run(build_server(fake_ctx).call_tool("teams_status", {"who": "Nobody Here"}))


def test_presence_label_and_parse() -> None:
    assert rules.presence_label("Busy", "InAMeeting") == "In a meeting"
    assert rules.presence_label("Available", "Available") == "Available"
    assert rules.presence_label("Busy", "InAConferenceCall") == "In a call"
    assert rules.presence_label("DoNotDisturb", "Presenting") == "Presenting"
    assert rules.presence_label("Offline", "SomeNewActivity") == "Some new activity"
    parsed = rules.parse_presence(
        {
            "availability": "Busy",
            "activity": "OutOfOffice",
            "calendarData": {"isOutOfOffice": True, "outOfOfficeNote": {"message": "<p>Back Mon</p>"}},
            "note": {"message": "<p>Focusing</p>"},
            "deviceType": "Mobile",
        }
    )
    assert parsed["status"] == "Out of office" and parsed["out_of_office"] is True
    assert parsed["ooo_note"] == "Back Mon" and parsed["message"] == "Focusing"
    assert parsed["available"] is False and parsed["device"] == "Mobile"


def test_fake_mark_read_and_message_actions(fake_ctx: AppContext) -> None:
    assert "*unread*" in call(fake_ctx, "teams_find")
    page = call(fake_ctx, "teams_read", chat="Jana")
    assert "Marked" in call(fake_ctx, "teams_mark_read", chat="Jana")
    jana = next(c for c in fake_ctx.fake.state.teams_chats if c.title == "Jana Nováková")
    assert jana.unread is False
    ref = next(
        lines[i - 1].split("]")[0].lstrip("[")
        for lines in [page.splitlines()]
        for i, line in enumerate(lines)
        if "NAT" in line
    )
    assert "Reacted like" in call(fake_ctx, "teams_update", message_ref=ref, react="thumbsup")
    msg = next(m for c in fake_ctx.fake.state.teams_chats for m in c.messages if m.id == "1002")
    assert msg.reactions.get("like") == 2
    assert "edited" in call(fake_ctx, "teams_update", message_ref=ref, edit="Approved.").lower()
    assert next(m for c in fake_ctx.fake.state.teams_chats for m in c.messages if m.id == "1002").edited
    assert "deleted" in call(fake_ctx, "teams_update", message_ref=ref, delete=True).lower()
    assert any(o.action == "teams_delete" for o in fake_ctx.fake.state.outbox)


def test_update_message_requires_one_action(fake_ctx: AppContext) -> None:
    page = call(fake_ctx, "teams_read", chat="Jana")
    ref = page.splitlines()[2].split("]")[0].lstrip("[")
    with pytest.raises(ToolError, match="exactly one"):
        asyncio.run(build_server(fake_ctx).call_tool("teams_update", {"message_ref": ref}))


def test_reaction_key() -> None:
    assert rules.reaction_key("thumbsup") == "like"
    assert rules.reaction_key("Heart") == "heart"
    assert rules.reaction_key("1f525_fire") == "1f525_fire"


def test_fake_digest_mentions_and_unread(fake_ctx: AppContext) -> None:
    out = call(fake_ctx, "teams_digest")
    assert "Mentions & replies to you" in out
    assert "Eva Horáková mentioned you in Platform squad" in out
    assert "Unread conversations" in out and "Jana Nováková" in out
    assert "<<<UNTRUSTED teams digest" in out
    ref = next(line.split("]")[0].lstrip("[") for line in out.splitlines() if "mentioned you" in line)
    assert ref.startswith("p:")
    preview = call(fake_ctx, "teams_draft", reply_to=ref, body="On it.")
    assert "NOT SENT" in preview


def test_parse_activity() -> None:
    import json as _json

    msg = {
        "content": "<p>hi</p>",
        "messagetype": "RichText/Html",
        "properties": {
            "isread": "false",
            "activity": _json.dumps(
                {
                    "activityType": "mention",
                    "sourceThreadId": "19:abc@thread.tacv2",
                    "sourceMessageId": "1790000000000",
                    "sourceUserImDisplayName": "Petr Svoboda",
                    "sourceThreadTopic": "Platform squad",
                    "messagePreview": "please review",
                    "activityTimestamp": 1790000000000,
                }
            ),
        },
    }
    item = rules.parse_activity(msg)
    assert item is not None
    assert item["who"] == "Petr Svoboda" and item["kind"] == "mentioned you"
    assert item["thread"] == "19:abc@thread.tacv2" and item["unread"] is True
    assert rules.parse_activity({"properties": {}}) is None


def test_fake_search(fake_ctx: AppContext) -> None:
    out = call(fake_ctx, "teams_search", query="NAT gateway")
    assert "destroy" in out and "<<<UNTRUSTED teams search" in out
    ref = next(line.split()[0] for line in out.splitlines() if line.startswith("c:"))
    assert "Terraform" in call(fake_ctx, "teams_read", chat=ref)
    assert "No Teams messages found" in call(fake_ctx, "teams_search", query="zzznomatch")


def test_search_kind_and_sender() -> None:
    assert rules.search_kind("Meeting") == "meeting"
    assert rules.search_kind("Channel") == "channel"
    assert rules.sender_name({"EmailAddress": {"Name": "Jana N", "Address": "j@x.com"}}) == "Jana N"
    assert rules.sender_name("Petr") == "Petr"
