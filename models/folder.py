from __future__ import annotations

import logging
from collections.abc import Iterable
from dataclasses import dataclass

from ..enums import FolderType, ItemType
from ..protocols import OlFolder
from ..utils import unpack_collection
from .mail_item import MailItem
from .node import ItemModel

logger = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class FolderListing:
    """Serializable folder tree entry for CLI and reporting."""

    path: str
    depth: int
    subfolder_count: int

    def as_row(self) -> tuple[object, ...]:
        return (self.path, self.depth, self.subfolder_count)


class Folder(ItemModel):
    """Represents an Outlook folder."""

    item_type = ItemType.FOLDER
    required_properties = ("Name", "Items", "Folders")
    inaccessible_error_message = "Provided Outlook item is not an accessible folder."

    def __init__(self, outlook_item: OlFolder) -> None:
        """Initialize Folder with an Outlook item."""
        super().__init__(outlook_item)
        self._ol_folder_item = outlook_item

    @property
    def name(self) -> str:
        """Return the folder name."""
        return self._ol_folder_item.Name

    @property
    def folder_type(self) -> str | None:
        """Return the folder type name based on the Outlook item type."""
        return FolderType.get_folder_type(self._ol_folder_item.Class) or "Undefined"

    @property
    def mail_items(self) -> list[MailItem]:
        """Return a list of MailItem objects in the folder."""
        return self.list_messages()

    @property
    def subfolders(self) -> list[Folder]:
        """Return a list of subfolders."""
        return list(self.iter_subfolders())

    def list_messages(
        self,
        limit: int | None = None,
        unread_only: bool = False,
    ) -> list[MailItem]:
        """Return mail items from this folder with optional filtering."""
        if limit is not None and limit < 1:
            raise ValueError("limit must be >= 1")

        mail_items = unpack_collection(self._ol_folder_item.Items, transformer=MailItem)

        if unread_only:
            mail_items = [item for item in mail_items if item.is_unread]

        if limit is not None:
            mail_items = mail_items[:limit]

        return mail_items

    def iter_subfolders(self) -> Iterable[Folder]:
        """Iterate over accessible child folders."""
        try:
            subfolders = unpack_collection(
                self._ol_folder_item.Folders, transformer=Folder
            )
            yield from subfolders
        except Exception:
            return

    def walk(
        self,
        recursive: bool = False,
        max_depth: int = 0,
        depth: int = 0,
        parent_path: str = "",
    ) -> list[FolderListing]:
        """Return folder tree entries rooted at this folder."""
        if recursive and max_depth < 0:
            raise ValueError("max_depth must be >= 0")

        path = f"{parent_path}/{self.name}" if parent_path else self.name
        subfolders = self.subfolders
        rows = [FolderListing(path=path, depth=depth, subfolder_count=len(subfolders))]

        if not recursive or depth >= max_depth:
            return rows

        for child in subfolders:
            rows.extend(
                child.walk(
                    recursive=recursive,
                    max_depth=max_depth,
                    depth=depth + 1,
                    parent_path=path,
                )
            )
        return rows

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

    def delete(self) -> None:
        """Delete this folder."""
        if FolderType.is_default_folder(self._ol_folder_item.Class):
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
