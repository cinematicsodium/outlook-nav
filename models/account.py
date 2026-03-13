from __future__ import annotations

from collections.abc import Iterable

from ..enums import ItemType
from ..models.folder import Folder
from ..protocols import OlAccount, OlFolder, OlStore
from ..utils import is_accessible_ol_item, unpack_collection


class Account:
    """Represents an Outlook account."""

    def __init__(self, ol_acct_item: OlAccount) -> None:
        """Initialize Account with an Outlook account item."""
        self.ol_account_item: OlAccount = ol_acct_item

    @classmethod
    def is_accessible_acct(cls, item: OlAccount) -> bool:
        """Check if the given item is an accessible Outlook account."""
        return is_accessible_ol_item(item=item, target_type=ItemType.ACCOUNT)

    @classmethod
    def from_outlook_item(cls, ol_acct_item: OlAccount) -> Account | None:
        """Create Account from an Outlook account item."""
        if not cls.is_accessible_acct(ol_acct_item):
            return None
        return cls(ol_acct_item)

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
    def root_folder(self) -> OlFolder:
        """Return the root folder of the account's default store."""
        return self.store.GetRootFolder()

    @property
    def folders(self) -> list[Folder]:
        """Return a list of folders in the account's root folder."""
        try:
            ol_folder_items: list[OlFolder] = unpack_collection(
                self.root_folder.Folders
            )
            if not ol_folder_items:
                return []
            folders: list[Folder | None] = [
                Folder.from_outlook_item(ol_folder) for ol_folder in ol_folder_items
            ]
            return [f for f in folders if f]
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

    def __str__(self) -> str:
        """Return the display name of the account."""
        return self.name

    def __repr__(self) -> str:
        """Return the string representation of the Account."""
        return f"Account({self.name})"
