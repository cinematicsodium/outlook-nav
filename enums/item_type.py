from enum import IntEnum


class ItemType(IntEnum):
    """Outlook object classes used by this package."""

    NAMESPACE = 1
    FOLDER = 2
    ADDRESS_ENTRY = 8
    MAIL_ITEM = 43
    ACCOUNT = 105
    STORE = 107
