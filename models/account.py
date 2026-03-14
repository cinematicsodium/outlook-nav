from __future__ import annotations

from collections.abc import Iterable

from ..enums import ItemType
from ..protocols import OlAccount, OlStore
from ..utils import unpack_collection
from .folder import Folder
from .node import ItemModel


class Account(ItemModel):
    """Represents an Outlook account."""

    item_type = ItemType.ACCOUNT
    required_properties = ("DisplayName", "SmtpAddress", "DeliveryStore")
    inaccessible_error_message = "Provided Outlook item is not an accessible account."

    def __init__(self, ol_acct_item: OlAccount) -> None:
        """Initialize Account with an Outlook account item."""
        super().__init__(ol_acct_item)
        self.ol_account_item: OlAccount = ol_acct_item

    @classmethod
    def is_accessible_acct(cls, item: OlAccount) -> bool:
        """Check if the given item is an accessible Outlook account."""
        return cls.is_accessible(item)

    @property
    def name(self) -> str:
        """Return the display name of the account."""
        return self.ol_account_item.DisplayName

    @property
    def address(self) -> str:
        """Return the email address of the account."""
        return self.ol_account_item.SmtpAddress

    @property
    def store(self) -> OlStore:
        """Return the default store associated with this account."""
        return self.ol_account_item.DeliveryStore

    @property
    def root_folder(self) -> Folder:
        """Return the root folder of the account's default store."""
        return Folder.from_outlook_item(self.store.GetRootFolder())

    @property
    def folders(self) -> list[Folder]:
        """Return a list of folders in the account's root folder."""
        try:
            folders = unpack_collection(self.root_folder.Folders, transformer=Folder)
            return folders
        except Exception:
            return []

    def get_folder(self, folder_name: str) -> Folder | None:
        """Return a folder by its name."""
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        target: str = folder_name.lower()
        for folder in self.folders:
            if (folder.name or "").lower() == target:
                return folder
        return None

    def find_folder(self, path: str | Iterable[str]) -> Folder | None:
        """Find a folder by path or iterable of folder names."""
        if isinstance(path, str):
            parts: list[str] = [
                segment.strip() for segment in path.split("/") if segment.strip()
            ]
        else:
            parts: list[str] = [
                str(segment).strip() for segment in path if str(segment).strip()
            ]

        if not parts:
            return None

        current: Folder | None = self.get_folder(parts[0])
        for segment in parts[1:]:
            if current is None:
                return None
            current = current.get_subfolder(segment)
        return current

    def matches(self, value: str | None) -> bool:
        """Return True when the value matches the account name or address."""
        if not value:
            return False
        target = value.lower()
        return target in {self.name.lower(), self.address.lower()}

    def __str__(self) -> str:
        """Return the display name of the account."""
        return self.name

    def __repr__(self) -> str:
        """Return the string representation of the Account."""
        return f"Account({self.name})"
