from enum import IntEnum


class FolderType(IntEnum):
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

    @classmethod
    def is_default_folder(cls, item_type: int) -> bool:
        """Check if the given item type corresponds to a default folder."""
        return cls.get_folder_type(item_type) is not None

    @classmethod
    def get_folder_type(cls, item_type: int) -> str | None:
        """Return the folder type name for a given Outlook item type."""
        try:
            return cls(item_type).name
        except ValueError:
            return None
