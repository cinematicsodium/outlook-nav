from enum import IntEnum


class FolderEnum(IntEnum):
    """Identifiers for Outlook's built-in default folders."""

    DELETED_ITEMS = 3
    OUTBOX = 4
    SENT_MAIL = 5
    INBOX = 6
    CALENDAR = 9
    CONTACTS = 10
    JOURNAL = 11
    NOTES = 12
    DRAFTS = 16
    CONFLICTS = 19
    LOCAL_FAILURES = 21
    JUNK = 23
    ARCHIVED_MAIL = 29


class ItemType(IntEnum):
    """Outlook object classes used by this package."""

    NAMESPACE = 1
    FOLDER = 2
    ADDRESS_ENTRY = 8
    MAIL_ITEM = 43
    ACCOUNT = 105
    STORE = 107
