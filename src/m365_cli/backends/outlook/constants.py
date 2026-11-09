from enum import IntEnum


class OlFolder(IntEnum):
    """Default folders (`OlDefaultFolders`)."""

    DELETED = 3
    OUTBOX = 4
    SENT = 5
    INBOX = 6
    CALENDAR = 9
    CONTACTS = 10
    DRAFTS = 16
    JUNK = 23


class OlItemType(IntEnum):
    """Item types for `Application.CreateItem` (`OlItemType`)."""

    MAIL = 0
    APPOINTMENT = 1


class OlObjectClass(IntEnum):
    """Item classes (`OlObjectClass`)."""

    APPOINTMENT = 26
    MAIL = 43
    REPORT = 46
    MEETING_REQUEST = 53
    MEETING_CANCELLATION = 54
    MEETING_RESPONSE_NEGATIVE = 55
    MEETING_RESPONSE_POSITIVE = 56
    MEETING_RESPONSE_TENTATIVE = 57


MEETING_ITEM_CLASSES = frozenset(
    range(OlObjectClass.MEETING_REQUEST, OlObjectClass.MEETING_RESPONSE_TENTATIVE + 1)
)


class OlMeetingStatus(IntEnum):
    """Meeting status of a calendar item (`OlMeetingStatus`)."""

    NON_MEETING = 0
    MEETING = 1
    RECEIVED = 3
    CANCELED = 5
    RECEIVED_AND_CANCELED = 7


class OlResponseStatus(IntEnum):
    """The user's response to a meeting (`OlResponseStatus`)."""

    NONE = 0
    ORGANIZED = 1
    TENTATIVE = 2
    ACCEPTED = 3
    DECLINED = 4
    NOT_RESPONDED = 5


class OlMeetingResponse(IntEnum):
    """Responses for `AppointmentItem.Respond` (`OlMeetingResponse`)."""

    TENTATIVE = 2
    ACCEPTED = 3
    DECLINED = 4


class OlRecipientType(IntEnum):
    """Recipient types (`OlMailRecipientType` / `OlMeetingRecipientType`).

    On meetings, TO is a required attendee, CC an optional attendee and BCC a resource (room).
    """

    TO = 1
    CC = 2
    BCC = 3


class OlRecurrenceType(IntEnum):
    """Recurrence types of a recurrence pattern (`OlRecurrenceType`)."""

    DAILY = 0
    WEEKLY = 1
    MONTHLY = 2


class OlFlagStatus(IntEnum):
    """Flag status of a mail (`OlFlagStatus`)."""

    NONE = 0
    COMPLETE = 1
    MARKED = 2


OL_DISCARD = 1
"""`OlInspectorClose.olDiscard`: close an item without saving it."""

OL_FORMAT_HTML = 2
"""`OlBodyFormat.olFormatHTML`: the item body is HTML."""

EXCHANGE_USER_TYPES = frozenset({0, 5})
"""`OlAddressEntryUserType` values of Exchange users.

olExchangeUserAddressEntry (0) and olExchangeRemoteUserAddressEntry (5).
"""

MAPI_E_NOT_FOUND = -2147221233
"""MAPI_E_NOT_FOUND (0x8004010F), surfaced through `pywintypes.com_error`."""

EXCHANGE_DELEGATE_STORE = 1
"""`OlExchangeStoreType.olExchangeMailbox`: a shared/delegate mailbox opened with the main account."""

EXCHANGE_PUBLIC_FOLDERS = 2
"""`OlExchangeStoreType.olExchangePublicFolder`: public folders, not a mailbox."""
