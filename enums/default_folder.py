from enum import IntEnum


class DefaultFolderEnum(IntEnum):
    DELETED_ITEMS = 3
    OUTBOX = 4
    SENT_MAIL = 5
    INBOX = 6
    CALENDAR = 9
    CONTACTS = 10
    JOURNAL = 11
    NOTES = 12
    CONFLICTS = 19
    LOCAL_FAILURES = 21
    JUNK = 23
    ARCHIVED_MAIL = 29
    DRAFTS = 16
