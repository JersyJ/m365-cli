from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pywintypes

from m365_cli.backends.outlook.constants import MEETING_ITEM_CLASSES, OlObjectClass
from m365_cli.backends.outlook.session import Com

_PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
_PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


@dataclass(frozen=True, slots=True)
class AttachmentInfo:
    """One attachment; `index` is 1-based, as in Outlook."""

    index: int
    name: str
    size: int
    inline: bool


def get(item: Com, attr: str, default: Any = None) -> Any:
    """Item property, or `default` when the item type does not have it."""
    try:
        return getattr(item, attr)
    except (AttributeError, pywintypes.com_error):
        return default


def kind_of(item: Com) -> str:
    """Item kind: mail | meeting | report | other."""
    cls = get(item, "Class", 0)
    if cls == OlObjectClass.MAIL:
        return "mail"
    if cls in MEETING_ITEM_CLASSES:
        return "meeting"
    return "report" if cls == OlObjectClass.REPORT else "other"


def attachments_of(item: Com) -> list[AttachmentInfo]:
    """Attachments of an item.

    An attachment counts as inline when it is hidden or has a content ID; these are usually images
    in the body, such as signature logos.
    """
    result: list[AttachmentInfo] = []
    for att in get(item, "Attachments", None) or []:
        inline = False
        try:
            accessor = att.PropertyAccessor
            inline = bool(accessor.GetProperty(_PR_ATTACHMENT_HIDDEN))
            inline = inline or bool(accessor.GetProperty(_PR_ATTACH_CONTENT_ID))
        except pywintypes.com_error:
            pass
        result.append(AttachmentInfo(index=att.Index, name=att.FileName, size=att.Size, inline=inline))
    return result


PROTECTED_MESSAGE_CLASSES = ("IPM.Note.SMIME", "IPM.Note.Microsoft.Voicemail.UM.CA")
"""Message classes whose content Outlook's object model cannot read."""


def is_protected(item: Com) -> bool:
    """True for rights-protected (MIP/IRM, Do Not Forward) or S/MIME-encrypted mails.

    Their body is only a wrapper ("This message is protected …"); the content is encrypted.
    """
    if get(item, "Permission", 0):
        return True
    if str(get(item, "MessageClass", "")).startswith(PROTECTED_MESSAGE_CLASSES[0]):
        return True
    return any(a.name.lower().endswith(".rpmsg") for a in attachments_of(item))


def split_names(value: str | None) -> list[str]:
    """Split an Outlook recipient string ("A; B") into names."""
    return [n.strip() for n in (value or "").split(";") if n.strip()]
