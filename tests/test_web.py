import json
from datetime import date, datetime

import pytest
from websockets.exceptions import ConnectionClosedError

from m365_cli.backends.browser import cdp as cdp_module
from m365_cli.backends.browser.cdp import ConnectionLostError
from m365_cli.backends.browser.session import HOOK_VERSION, BrowserConfig, PageSession
from m365_cli.backends.outlook_web.session import ews_id, from_utc, from_wall, quote, rest_id, to_utc, wall
from m365_cli.core.context import AppContext
from m365_cli.core.errors import BackendUnavailableError, InvalidInputError
from m365_cli.features.calendar.api import calendar_api
from m365_cli.features.calendar.models import Repeat
from m365_cli.features.calendar.web import _recurrence, _set_field
from m365_cli.features.mail.models import MailQuery, MailSummary
from m365_cli.features.mail.render import summary_line
from m365_cli.features.mail.web import (
    _CONTENT_CLASS,
    _MESSAGE_CLASS,
    _above_original,
    _filter,
    _inline_images,
    _is_protected,
    _kind,
    _matches,
    _scheduled,
)


def _props(**values: str) -> dict:
    ids = {"cls": _MESSAGE_CLASS, "cc": _CONTENT_CLASS}
    return {"SingleValueExtendedProperties": [{"PropertyId": ids[k], "Value": v} for k, v in values.items()]}


def test_ids_round_trip() -> None:
    ews = "AAkALg/AAAA+AAA="
    assert rest_id(ews) == "AAkALg-AAAA_AAA="
    assert ews_id(rest_id(ews)) == ews


def test_times_round_trip_and_wall_clock() -> None:
    local = datetime(2026, 11, 17, 9, 30)
    assert from_utc(to_utc(local)) == local
    assert from_utc(None) is None
    assert from_wall({"DateTime": "2026-11-18T16:00:00.0000000", "TimeZone": "UTC"}) == datetime(
        2026, 11, 18, 16
    )
    assert wall(local, "Central Europe Standard Time") == {
        "DateTime": "2026-11-17T09:30:00",
        "TimeZone": "Central Europe Standard Time",
    }


def test_filter_puts_sort_field_first_and_quotes() -> None:
    assert _filter(MailQuery()) == ""
    flt = _filter(MailQuery(unread_only=True, subject="O'Neil", sender="jan", flagged=True))
    assert flt.startswith("ReceivedDateTime ge 1900-01-01T00:00:00Z and ")
    assert "IsRead eq false" in flt
    assert "contains(Subject,'O''Neil')" in flt
    assert "contains(From/EmailAddress/Name,'jan')" in flt
    assert "Flag/FlagStatus eq 'Flagged'" in flt
    assert quote("a'b") == "a''b"


def test_search_results_get_the_other_filters() -> None:
    m = {
        "ReceivedDateTime": "2026-11-17T08:00:00Z",
        "IsRead": False,
        "HasAttachments": True,
        "Flag": {"FlagStatus": "NotFlagged"},
        "Subject": "Re: GitHub Copilot Access",
        "From": {"EmailAddress": {"Name": "Niklas", "Address": "niklas@example.com"}},
    }
    assert _matches(m, MailQuery(text="x", unread_only=True, has_attachments=True, sender="niklas"))
    assert not _matches(m, MailQuery(text="x", flagged=True))
    assert not _matches(m, MailQuery(text="x", subject="invoice"))
    assert not _matches(m, MailQuery(text="x", since=datetime(2030, 1, 1)))


def test_kind_and_protected() -> None:
    assert _kind({}) == "mail"
    assert _kind({"@odata.type": "#Microsoft.OutlookServices.EventMessageRequest"}) == "meeting"
    assert _kind(_props(cls="REPORT.IPM.Note.NDR")) == "report"
    server_spelling = {
        "SingleValueExtendedProperties": [{"PropertyId": "String 0x1a", "Value": "REPORT.IPM.Note.NDR"}]
    }
    assert _kind(server_spelling) == "report"
    assert _kind(_props(cls="IPM.Task")) == "other"
    assert _is_protected(_props(cls="IPM.Note", cc="rpmsg.message"))
    assert _is_protected(_props(cls="IPM.Note.SMIME"))
    assert not _is_protected(_props(cls="IPM.Note"))


def test_inline_images_become_cid_attachments() -> None:
    html = '<p>Hi</p><img src="data:image/png;base64,iVBOR\nw0K"><img src="data:image/jpg;base64,AAAA">'
    fragment, images = _inline_images(html)
    assert "data:" not in fragment
    assert 'src="cid:m365img1@m365-cli"' in fragment
    assert images == [("m365img1@m365-cli", "png", "iVBORw0K"), ("m365img2@m365-cli", "jpeg", "AAAA")]


def test_reply_text_goes_above_the_original() -> None:
    draft = {"Body": {"Content": '<html><body dir="ltr"><div>original</div></body></html>'}}
    assert _above_original(draft, "<div>ours</div>") == (
        '<html><body dir="ltr"><div>ours</div><br><div>original</div></body></html>'
    )
    assert _above_original({"Body": {"Content": "plain"}}, "<b>x</b>").count("<body>") == 1


def test_recurrence_and_location_field() -> None:
    zone = "Central Europe Standard Time"
    weekly = _recurrence(Repeat(every="week", count=4), ["mon"], datetime(2027, 3, 15, 9), zone)
    assert weekly["Pattern"] == {
        "Type": "Weekly",
        "Interval": 1,
        "DaysOfWeek": ["Monday"],
        "FirstDayOfWeek": "Monday",
    }
    assert weekly["Range"] == {
        "StartDate": "2027-03-15",
        "RecurrenceTimeZone": zone,
        "Type": "Numbered",
        "NumberOfOccurrences": 4,
    }
    monthly = _recurrence(Repeat(every="month", until=date(2027, 6, 1)), [], datetime(2027, 3, 15, 9), zone)
    assert monthly["Pattern"]["DayOfMonth"] == 15
    assert monthly["Range"]["Type"] == "EndDate"
    assert _set_field("Location", "Room 1")["Path"]["FieldURI"] == "EnhancedLocation"
    assert _set_field("Subject", "x")["Item"] == {"__type": "CalendarItem:#Exchange", "Subject": "x"}


def test_series_reply_needs_a_recurring_meeting(fake_ctx: AppContext) -> None:
    cal = calendar_api(fake_ctx)
    events = {e.subject: e for e in cal.find(datetime(2026, 11, 9), datetime(2026, 11, 12)).items}
    assert "(whole series)" in cal.respond(events["Daily standup"].ref, "decline", series=True)
    with pytest.raises(InvalidInputError):
        cal.respond(events["Weekly GitLab Meeting"].ref, "accept", series=True)


def test_scheduled_mark_only_for_waiting_drafts() -> None:
    later = {
        "IsDraft": True,
        "SingleValueExtendedProperties": [
            {"PropertyId": "SystemTime 0x3fef", "Value": "2099-01-01T08:00:00Z"}
        ],
    }
    assert _scheduled(later) is not None
    assert _scheduled({**later, "IsDraft": False}) is None
    assert _scheduled({"IsDraft": True}) is None
    summary = MailSummary(
        ref="m:AAAAAAA", received=None, sender="me", subject="s", scheduled=datetime(2099, 1, 1, 9)
    )
    assert "[scheduled " in summary_line(summary)


class _Ws:
    """A fake DevTools websocket: answers every call, or fails as `fail` says."""

    def __init__(self, fail: str = "") -> None:
        self.fail = fail
        self.sent: list[dict] = []

    def send(self, text: str) -> None:
        if self.fail == "closed":
            raise ConnectionClosedError(None, None)
        self.sent.append(json.loads(text))

    def recv(self, timeout: float | None = None) -> str:
        if self.fail == "timeout":
            raise TimeoutError
        message = self.sent[-1]
        if message["method"] == "Browser.getWindowForTarget":
            return json.dumps({"id": message["id"], "result": {"windowId": 1}})
        ready = "status()" in message["params"].get("expression", "")
        value = {"version": HOOK_VERSION, "ok": True} if ready else {"status": 200}
        return json.dumps({"id": message["id"], "result": {"result": {"value": value}}})

    def close(self) -> None:
        pass


class _App(PageSession):
    url = "https://app.example/"
    hosts = ("app.example",)

    def _is_ready(self, state: dict) -> bool:
        return bool(state.get("ok"))


def _app(monkeypatch: pytest.MonkeyPatch, tmp_path, sockets: list[_Ws]) -> tuple[_App, list[int]]:
    """A session already connected to `sockets[0]`, with a fake browser that is not running."""
    monkeypatch.setattr(cdp_module, "connect", lambda *a, **k: sockets.pop(0))
    app = _App(BrowserConfig(port=1, profile_dir=tmp_path, window="minimized"))
    started: list[int] = []
    tab = {"url": "https://app.example/", "webSocketDebuggerUrl": "ws://tab"}
    monkeypatch.setattr(app, "_tabs", lambda: [tab] if started else None)
    monkeypatch.setattr(app, "_start_browser", lambda: started.append(1))
    monkeypatch.setattr(app, "_wait_for_tabs", lambda: [tab])
    app._cdp = cdp_module.CdpConnection("ws://tab")
    return app, started


def test_lost_browser_connection_starts_the_browser_again(monkeypatch, tmp_path):
    fresh = _Ws()
    app, started = _app(monkeypatch, tmp_path, [_Ws("closed"), fresh])
    dead = app._cdp
    assert app._call({}) == {"status": 200}
    assert started == [1]
    assert app._cdp is not dead
    minimized = {"windowId": 1, "bounds": {"windowState": "minimized"}}
    assert any(m["method"] == "Browser.setWindowBounds" and m["params"] == minimized for m in fresh.sent)


def test_slow_browser_call_is_not_retried(monkeypatch, tmp_path):
    sockets = [_Ws("timeout"), _Ws()]
    app, started = _app(monkeypatch, tmp_path, sockets)
    slow = app._cdp
    with pytest.raises(BackendUnavailableError) as raised:
        app._call({})
    assert not isinstance(raised.value, ConnectionLostError)
    assert app._cdp is slow
    assert started == []
    assert len(sockets) == 1
