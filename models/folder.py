from __future__ import annotations

import logging
from collections.abc import Iterable
from dataclasses import dataclass

from ..enums import FolderType, ItemType
from ..models.mail_item import MailItem
from ..protocols import OlFolder
from ..utils import is_accessible_ol_item

logger = logging.getLogger(__name__)


@dataclass
class Folder:
    """Represents an Outlook folder."""

    def __init__(self, outlook_item: OlFolder) -> None:
        """Initialize Folder with an Outlook item."""
        self._ol_folder_item: OlFolder = outlook_item
        if not self.is_folder_accessible(outlook_item):
            raise ValueError("Provided Outlook item is not an accessible folder.")

    @classmethod
    def from_outlook_item(cls, item: object) -> Folder | None:
        """Create Folder from an Outlook item if accessible."""
        if not cls.is_folder_accessible(item):
            return None
        return cls(item)

    @classmethod
    def is_folder_accessible(cls, item: object) -> bool:
        """Check if the Outlook folder item is accessible."""
        return is_accessible_ol_item(
            item=item,
            target_type=ItemType.FOLDER,
            properties=cls.interface_properties(),
        )

    @classmethod
    def interface_properties(cls) -> tuple[str, str, str]:
        """Return interface properties for Outlook folder."""
        return ("Name", "Items", "Folders")

    @staticmethod
    def _is_default_folder(item: OlFolder) -> bool:
        """Check if the folder is a default folder."""
        try:
            return bool(FolderType(item.Class))
        except Exception:
            return False

    @property
    def name(self) -> str:
        """Return the folder name."""
        return self._ol_folder_item.Name

    @property
    def mail_items(self) -> list[MailItem]:
        """Return a list of MailItem objects in the folder."""
        folder_items = self._ol_folder_item.Items
        item_count: int = folder_items.Count

        if not (folder_items and item_count):
            return []

        outlook_items: list[object] = [
            folder_items.Item(idx) for idx in range(1, item_count + 1)
        ]
        mail_items: list[MailItem | None] = [
            MailItem.from_outlook_item(item) for item in outlook_items
        ]
        valid_items: list[MailItem] = [item for item in mail_items if item]
        return valid_items

    @property
    def subfolders(self) -> list[Folder]:
        """Return a list of subfolders."""
        try:
            folders = self._ol_folder_item.Folders
            count: int = folders.Count

            if not (folders and count):
                return []

            outlook_items: list[object] = [
                folders.Item(idx) for idx in range(1, count + 1)
            ]

            folder_items: list[Folder | None] = [
                Folder.from_outlook_item(folder)
                for folder in outlook_items
                if folder.Class == ItemType.FOLDER
            ]

            return [folder for folder in folder_items if folder]

        except Exception:
            return []

    def get_item(self, index: int) -> MailItem | None:
        """Get MailItem at the specified index."""
        if index < 0:
            raise ValueError("index must be >= 0")

        items: list[MailItem] = self.mail_items
        if index >= len(items):
            return None
        return items[index]

    def get_subfolder(self, folder_name: str) -> Folder | None:
        """Get subfolder by name."""
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")

        target: str = folder_name.lower()
        for folder in self.subfolders:
            if (folder.name or "").lower() == target:
                return folder
        return None

    def create_subfolder(self, folder_name: str) -> Folder | None:
        """Create a new subfolder with the given name."""
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        if not folder_name.strip():
            raise ValueError("folder_name cannot be empty")
        try:
            folders = self._ol_folder_item.Folders
            if folders is None:
                raise AttributeError("Folder does not expose a Folders collection")
            created = folders.Add(folder_name)
            return Folder(created)
        except Exception:
            logger.error(
                "Error creating subfolder '%s' in '%s'", folder_name, self.name
            )
            return None

    def delete_subfolder(self, folder_name: str) -> bool:
        """Delete a subfolder by name."""
        folder: Folder | None = self.get_subfolder(folder_name)
        if folder is None:
            return False
        folder.delete()
        return True

    def move_item(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        """Move a MailItem to another Folder."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        if not isinstance(destination, Folder):
            raise ValueError("destination must be a Folder")
        return mail_item.move_to(destination)

    def delete_item(self, mail_item: MailItem) -> None:
        """Delete a MailItem from the folder."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def extend_items(self, values: Iterable[MailItem]) -> list[MailItem]:
        """Extend mail items with values."""
        return list(values)

    def delete(self) -> None:
        """Delete this folder."""
        if self._is_default_folder(self._ol_folder_item):
            raise ValueError(f"Cannot delete default folder: {self.folder_type}")
        try:
            self._ol_folder_item.Delete()
        except Exception:
            logger.error("Error deleting folder '%s'", self.name)

    def __str__(self) -> str:
        """Return string representation of the folder."""
        return self.name

    def __repr__(self) -> str:
        """Return repr string for the folder."""
        return f"Folder(name={self.name})"
