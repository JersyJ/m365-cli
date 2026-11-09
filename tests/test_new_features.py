import zipfile
from datetime import datetime
from pathlib import Path

import pytest

from m365_cli.core.context import AppContext
from m365_cli.core.doctext import extract_text
from m365_cli.core.errors import InvalidInputError, NotFoundError
from m365_cli.features.calendar.api import calendar_api
from m365_cli.features.mail.api import mail_api
from m365_cli.features.mail.models import MailQuery
from m365_cli.features.mail.rules import normalize_subject, pick_attachment
from tests.test_mcp_server import call

TOMORROW = (datetime(2026, 11, 10), datetime(2026, 11, 11))


def test_normalize_subject_strips_reply_prefixes() -> None:
    assert normalize_subject("RE: Fw: AW:  Budget") == "Budget"
    assert normalize_subject("Re[2]: SV: Plan") == "Plan"
    assert normalize_subject("Retro meeting") == "Retro meeting"


def test_pick_attachment_by_index_name_or_part() -> None:
    options = [(1, "invoice.pdf"), (2, "logo.png")]
    assert pick_attachment(options, "2") == (2, "logo.png")
    assert pick_attachment(options, "INVOICE.PDF") == (1, "invoice.pdf")
    assert pick_attachment(options, "logo") == (2, "logo.png")
    with pytest.raises(NotFoundError, match=r"Attachments: 1\. invoice\.pdf"):
        pick_attachment(options, "contract")


def test_digest_groups_by_sender(fake_ctx: AppContext) -> None:
    groups = mail_api(fake_ctx).digest(MailQuery(since=datetime(2026, 11, 8)))
    renovate = next(g for g in groups if g.sender == "renovate[bot]")
    assert (renovate.count, renovate.unread) == (3, 3)
    assert renovate.newest_ref.startswith("m:")
    assert groups[0].count >= groups[-1].count


def test_bulk_update_by_filter_in_one_call(fake_ctx: AppContext) -> None:
    result = call(fake_ctx, "mail_update", where={"sender": "renovate", "unread_only": True}, read=True)
    assert result.startswith("3 mail(s): marked read")
    assert all(not m.unread for m in fake_ctx.fake.state.mails if m.sender_name == "renovate[bot]")
    assert "No mails match" in call(
        fake_ctx, "mail_update", where={"sender": "renovate", "unread_only": True}, read=True
    )


def test_update_needs_exactly_one_target(fake_ctx: AppContext) -> None:
    from mcp.server.mcpserver.exceptions import ToolError

    with pytest.raises(ToolError, match="either mail_refs or where"):
        call(fake_ctx, "mail_update", read=True)


def test_find_all_folders_shows_folder(fake_ctx: AppContext) -> None:
    page = mail_api(fake_ctx).find(MailQuery(folder="all", subject="poison queue"))
    assert [m.folder for m in page.items] == ["Inbox/Alerts"]
    assert "[Inbox/Alerts]" in call(fake_ctx, "mail_find", folder="all", subject="poison")


def test_thread_and_waiting(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    sizing = mail.find(MailQuery(folder="sent", subject="runner sizing")).items[0]
    thread = mail.thread(sizing.ref)
    assert [m.subject for m in thread] == ["GitLab runner sizing", "RE: GitLab runner sizing"]
    assert thread[1].quoted_cut
    waiting = mail.waiting(days=3, within_days=14)
    assert [w.subject for w in waiting] == ["Budget numbers for Q1"]
    assert waiting[0].days == 5


def test_attachment_text_and_scheduled_send(fake_ctx: AppContext) -> None:
    mail = mail_api(fake_ctx)
    invoice = mail.find(MailQuery(sender="billing")).items[0]
    text = mail.attachment_text(invoice.ref, "1")
    assert "IBAN: CZ65" in text.text and text.name == "invoice-2026-114.pdf"
    draft = mail.compose("new", to=["Jana"], subject="Status", body="Done.")
    with pytest.raises(InvalidInputError, match="in the past"):
        mail.send(draft.ref, send_at=datetime(2026, 11, 1))
    assert mail.send(draft.ref, send_at=datetime(2026, 11, 10, 8)).startswith("Scheduled")
    assert fake_ctx.fake.state.outbox[0].send_at == datetime(2026, 11, 10, 8)


def test_calendar_update_keeps_duration_and_needs_send(fake_ctx: AppContext) -> None:
    cal = calendar_api(fake_ctx)
    budget = next(e for e in cal.find(*TOMORROW).items if e.subject == "Budget planning")
    moved = cal.update(budget.ref, start=datetime(2026, 11, 10, 16), add_required=["Jana"])
    assert (moved.start.hour, moved.end.hour) == (16, 17) and moved.is_update
    assert moved.required == ["Eva Horáková", "Jana Nováková"]
    event = next(e for e in fake_ctx.fake.state.events if e.subject == "Budget planning")
    assert not event.sent and fake_ctx.fake.state.outbox == []
    gitlab = next(e for e in cal.find(*TOMORROW).items if e.subject == "Weekly GitLab Meeting")
    with pytest.raises(InvalidInputError, match="not the organizer"):
        cal.update(gitlab.ref, location="Room 1")


def test_calendar_cleanup_previews_then_removes(fake_ctx: AppContext) -> None:
    preview = call(fake_ctx, "cal_cleanup", start="today", days=7)
    assert "Canceled: Old sync" in preview and "Removed" not in preview
    assert any(e.canceled for e in fake_ctx.fake.state.events)
    assert "Removed 1" in call(fake_ctx, "cal_cleanup", start="today", days=7, apply=True)
    assert not any(e.canceled for e in fake_ctx.fake.state.events)


def _zip(path: Path, files: dict[str, str]) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        for name, content in files.items():
            archive.writestr(name, content)
    return path


def test_extract_text_from_office_files(tmp_path: Path) -> None:
    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    docx = _zip(
        tmp_path / "a.docx",
        {
            "word/document.xml": f"<w:document {w}><w:body><w:p><w:r><w:t>Hello</w:t>"
            "</w:r><w:r><w:t> world</w:t></w:r></w:p></w:body>"
            "</w:document>"
        },
    )
    a = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    pptx = _zip(tmp_path / "b.pptx", {"ppt/slides/slide1.xml": f"<p {a}><a:t>Agenda</a:t></p>"})
    s = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    xlsx = _zip(
        tmp_path / "c.xlsx",
        {
            "xl/sharedStrings.xml": f"<sst {s}><si><t>Total</t></si></sst>",
            "xl/worksheets/sheet1.xml": f'<worksheet {s}><sheetData><row><c t="s"><v>0</v></c>'
            "<c><v>42</v></c></row>"
            "</sheetData></worksheet>",
        },
    )
    (tmp_path / "d.txt").write_text("plain", encoding="utf-8")
    assert extract_text(docx) == "Hello world"
    assert extract_text(pptx) == "[Slide 1] Agenda"
    assert extract_text(xlsx) == "[sheet1]\nTotal\t42"
    assert extract_text(tmp_path / "d.txt") == "plain"
    with pytest.raises(InvalidInputError, match="supported are"):
        extract_text(tmp_path / "e.zip")
