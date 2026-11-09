from __future__ import annotations

from m365_cli.core.paging import Page, more_hint
from m365_cli.core.text import one_line, untrusted
from m365_cli.core.timeutil import fmt
from m365_cli.features.mail.models import Draft, MailMessage, MailSummary


def summary_line(m: MailSummary) -> str:
    """One line per mail: ref, '*' if unread, time, sender, subject, [att]/[flag]; preview below."""
    marks = "".join(mark for cond, mark in ((m.has_attachments, " [att]"), (m.flagged, " [flag]")) if cond)
    kind = f" {{{m.kind}}}" if m.kind != "mail" else ""
    unread = "*" if m.unread else " "
    sender, subject = one_line(m.sender, 28), one_line(m.subject, 80)
    line = f"{m.ref} {unread} {fmt(m.received)} | {sender} | {subject}{marks}{kind}"
    return f"{line}\n    > {m.preview}" if m.preview else line


def summaries(page: Page[MailSummary]) -> str:
    """All lines of a page, plus the hint for the next page if there is one."""
    if not page.items:
        return "(no messages)"
    lines = [summary_line(m) for m in page.items]
    hint = more_hint(page)
    return "\n".join([*lines, hint] if hint else lines)


def message(m: MailMessage, *, mark_untrusted: bool = False) -> str:
    """Header lines and body of a mail. `mark_untrusted` wraps the body in markers (for the model)."""
    lines = [
        f"{m.ref} | {fmt(m.received)} | {one_line(m.subject, 200)}",
        f"From: {m.sender}" + (f" <{m.sender_email}>" if m.sender_email else ""),
        f"To: {m.to}",
    ]
    if m.cc:
        lines.append(f"Cc: {m.cc}")
    if m.categories:
        lines.append(f"Categories: {m.categories}")
    files = [a for a in m.attachments if not a.inline]
    if files:
        lines.append("Attachments: " + ", ".join(f"{a.name} ({max(1, a.size // 1024)} KB)" for a in files))
    body = m.body or "(empty body)"
    lines += ["", untrusted("MAIL BODY", body) if mark_untrusted else body]
    if m.quoted_cut or m.truncated:
        omitted = "quoted history" if m.quoted_cut else "rest of the body"
        lines.append(f"[{omitted} omitted; read again with response_format='detailed']")
    return "\n".join(lines)


def draft(d: Draft) -> str:
    """Preview of a draft."""
    lines = [f"DRAFT {d.ref} ({d.mode.replace('_', '-')}) - NOT SENT", "To: " + "; ".join(d.to)]
    if d.cc:
        lines.append("Cc: " + "; ".join(d.cc))
    if d.bcc:
        lines.append("Bcc: " + "; ".join(d.bcc))
    lines.append(f"Subject: {d.subject}")
    if d.attachments:
        lines.append("Attachments: " + ", ".join(d.attachments))
    lines += ["", d.text]
    if d.includes_original:
        lines.append("[+ original message quoted below]")
    return "\n".join(lines)
