from __future__ import annotations

from datetime import timedelta
from typing import TYPE_CHECKING

from m365_cli.core.errors import M365Error
from m365_cli.core.timeutil import fmt, now

if TYPE_CHECKING:
    from m365_cli.core.context import AppContext

_UNAVAILABLE = "{title}: unavailable ({reason})"


def briefing(ctx: AppContext) -> str:
    """A cross-feature 'what needs me today' summary: meetings, mail and Teams.

    Each section is skipped when its feature is disabled and degrades to a short note when its backend
    is unavailable, so one failure never hides the rest.
    """
    disabled = set(ctx.settings.disabled_features)
    moment = now()
    sections = [f"Briefing for {fmt(moment)}"]
    if "calendar" not in disabled:
        sections.append(_calendar(ctx, moment))
    if "mail" not in disabled:
        sections.append(_mail(ctx, moment))
    if "teams" not in disabled:
        sections.append(_teams(ctx))
    return "\n\n".join(s for s in sections if s)


def _calendar(ctx: AppContext, moment: object) -> str:
    from datetime import datetime

    from m365_cli.features.calendar.api import calendar_api

    assert isinstance(moment, datetime)
    start = moment.replace(hour=0, minute=0, second=0, microsecond=0)
    try:
        page = calendar_api(ctx).find(start, start + timedelta(days=1), limit=20)
    except M365Error as exc:
        return _UNAVAILABLE.format(title="Today's meetings", reason=exc)
    upcoming = [e for e in page.items if e.all_day or e.end >= moment]
    if not upcoming:
        return "Today's meetings: none left today."
    lines = [f"Today's meetings ({len(upcoming)}):"]
    lines += [f"  {fmt(e.start)}-{fmt(e.end, base=e.start)} {e.subject}" for e in upcoming]
    return "\n".join(lines)


def _mail(ctx: AppContext, moment: object) -> str:
    from datetime import datetime

    from m365_cli.features.mail.api import mail_api
    from m365_cli.features.mail.models import MailQuery

    assert isinstance(moment, datetime)
    try:
        api = mail_api(ctx)
        groups = api.digest(MailQuery(unread_only=True, since=moment - timedelta(days=2)), max_items=200)
        waiting = api.waiting(days=3, limit=5)
    except M365Error as exc:
        return _UNAVAILABLE.format(title="Mail", reason=exc)
    total = sum(g.count for g in groups)
    lines = [f"Unread mail ({total} from {len(groups)} sender(s)):"]
    lines += [f"  {g.sender}: {g.count} - {g.top_subjects[0][0]}" for g in groups[:5] if g.top_subjects]
    if waiting:
        lines.append(f"Waiting on a reply ({len(waiting)}):")
        lines += [f"  {w.subject} -> {w.to} ({w.days}d)" for w in waiting[:5]]
    return "\n".join(lines)


def _teams(ctx: AppContext) -> str:
    from m365_cli.features.teams import render
    from m365_cli.features.teams.api import teams_api

    try:
        digest = teams_api(ctx).digest(unread_only=True, limit=10)
    except M365Error as exc:
        return _UNAVAILABLE.format(title="Teams", reason=exc)
    if not digest.mentions and not digest.unread:
        return "Teams: nothing unread and no new mentions."
    return render.digest(digest)
