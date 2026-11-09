from __future__ import annotations

from m365_cli.core.paging import Page, more_hint
from m365_cli.core.text import one_line, untrusted
from m365_cli.core.timeutil import fmt
from m365_cli.features.mail.models import (
    AttachmentText,
    DigestGroup,
    Draft,
    MailMessage,
    MailSummary,
    WaitingItem,
)


def summary_line(m: MailSummary) -> str:
    """One line per mail: ref, '*' if unread, time, sender, subject, [att]/[flag]; preview below."""
    flags = ((m.has_attachments, " [att]"), (m.flagged, " [flag]"), (m.protected, " [protected]"))
    marks = "".join(mark for cond, mark in flags if cond)
    kind = f" {{{m.kind}}}" if m.kind != "mail" else ""
    unread = "*" if m.unread else " "
    sender, subject = one_line(m.sender, 28), one_line(m.subject, 80)
    folder = f" [{m.folder}]" if m.folder else ""
    line = f"{m.ref} {unread} {fmt(m.received)} | {sender} | {subject}{marks}{kind}{folder}"
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
    if m.protected:
        lines += ["", PROTECTED_NOTE]
        return "\n".join(lines)
    body = m.body or "(empty body)"
    lines += ["", untrusted("MAIL BODY", body) if mark_untrusted else body]
    if m.quoted_cut or m.truncated:
        omitted = "quoted history" if m.quoted_cut else "rest of the body"
        lines.append(f"[{omitted} omitted; read again with response_format='detailed']")
    return "\n".join(lines)


PROTECTED_NOTE = (
    "[Protected mail (Microsoft Information Protection / encryption): the content is encrypted and can "
    "only be read in Outlook. Tell the user to open it there; do not guess its content.]"
)


def draft(d: Draft) -> str:
    """Preview of a draft."""
    lines = [f"DRAFT {d.ref} ({d.mode.replace('_', '-')}) - NOT SENT"]
    if d.from_mailbox:
        lines.append(f"From: {d.from_mailbox}")
    lines.append("To: " + "; ".join(d.to))
    if d.cc:
        lines.append("Cc: " + "; ".join(d.cc))
    if d.bcc:
        lines.append("Bcc: " + "; ".join(d.bcc))
    lines.append(f"Subject: {d.subject}")
    if d.attachments:
        lines.append("Attachments: " + ", ".join(d.attachments))
    lines += ["", d.text]
    if d.signature:
        lines.append(f"[+ signature: {d.signature}]")
    if d.includes_original:
        lines.append("[+ original message quoted below]")
    return "\n".join(lines)


def digest(groups: list[DigestGroup]) -> str:
    """One line per sender: count, unread, newest time, top subjects and the newest mail's ref."""
    if not groups:
        return "(no messages)"
    lines = []
    for g in groups:
        subjects = "; ".join(
            f"{one_line(s, 60)} ({n})" if n > 1 else one_line(s, 60) for s, n in g.top_subjects
        )
        newest = f"newest {fmt(g.newest)} {g.newest_ref}"
        lines.append(f"{g.count:3} ({g.unread} unread) | {one_line(g.sender, 30)} | {newest} | {subjects}")
    return "\n".join(lines)


def thread(messages: list[MailMessage], *, mark_untrusted: bool = False) -> str:
    """All messages of a conversation, oldest first, each with a short header."""
    if not messages:
        return "(empty conversation)"
    parts = [f"Conversation: {one_line(messages[0].subject, 120)} ({len(messages)} messages)"]
    for m in messages:
        header = f"--- {m.ref} | {fmt(m.received)} | {one_line(m.sender, 40)}"
        if m.protected:
            parts += [header, PROTECTED_NOTE]
            continue
        body = m.body or "(empty body)"
        parts += [header, untrusted("MAIL BODY", body) if mark_untrusted else body]
    return "\n".join(parts)


def waiting(items: list[WaitingItem]) -> str:
    """One line per sent mail still waiting for a reply."""
    if not items:
        return "(no sent mails waiting for a reply)"
    return "\n".join(
        f"{w.ref} | sent {fmt(w.sent)} ({w.days} days) | to {one_line(w.to, 40)} | {one_line(w.subject, 80)}"
        for w in items
    )


def attachment_text(a: AttachmentText, *, mark_untrusted: bool = False) -> str:
    """Extracted attachment text with its name."""
    body = a.text + ("\n[… cut]" if a.truncated else "")
    return f"Attachment: {a.name}\n" + (untrusted("ATTACHMENT", body) if mark_untrusted else body)
