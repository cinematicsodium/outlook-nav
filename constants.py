import re
from enum import IntEnum


class FolderEnum(IntEnum):
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
    MANAGED_EMAIL = 29
    DRAFTS = 16


EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")

DIGIT_REGEX = re.compile(r"[0-9]+")
