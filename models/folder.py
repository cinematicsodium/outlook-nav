from __future__ import annotations
import logging
from collections.abc import Iterable
from dataclasses import dataclass
from functools import cached_property
from ..enums import ItemType
from ..protocols import OlFolder
from ..utils import unpack_collection
from .base import ItemModel
from .mail_item import MailItem

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

    def __init__(self, item: OlFolder) -> None:
        """Initialize Folder with an Outlook item."""
        super().__init__(item)
        self._ol_folder_item = item

    @cached_property
    def name(self) -> str:
        """Return the folder name."""
        return self._ol_folder_item.Name

    @cached_property
    def folder_path(self) -> str:
        return self._ol_folder_item.FolderPath

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
        items = self._ol_folder_item.Items
        sort = getattr(items, "Sort", None)
        if callable(sort):
            try:
                sort("[ReceivedTime]", True)
            except Exception:
                logger.debug("Unable to sort messages in '%s'", self.name)
        if unread_only:
            restrict = getattr(items, "Restrict", None)
            if callable(restrict):
                try:
                    items = restrict("[UnRead] = True")
                    unread_only = False
                except Exception:
                    logger.debug("Unable to filter unread messages in '%s'", self.name)
        return unpack_collection(
            items,  # type: ignore
            transformer=MailItem,
            limit=limit,
            predicate=(lambda item: bool(getattr(item, "UnRead", False)))
            if unread_only
            else None,
        )  # type: ignore

    def iter_subfolders(self) -> Iterable[Folder]:
        """Iterate over accessible child folders."""
        try:
            subfolders = unpack_collection(
                self._ol_folder_item.Folders, transformer=Folder
            )
            yield from subfolders
        except Exception:
            logger.warning("Unable to enumerate subfolders for '%s'", self.name)
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
        folder_listing_entries = [
            FolderListing(path=path, depth=depth, subfolder_count=len(subfolders))
        ]
        if not recursive or depth >= max_depth:
            return folder_listing_entries
        for subfolder in subfolders:
            folder_listing_entries.extend(
                subfolder.walk(
                    recursive=recursive,
                    max_depth=max_depth,
                    depth=depth + 1,
                    parent_path=path,
                )
            )
        return folder_listing_entries

    def get_item(self, index: int) -> MailItem | None:
        """Get MailItem at the specified index."""
        if index < 0:
            raise ValueError("index must be >= 0")
        items = self._ol_folder_item.Items
        if index >= items.Count:
            return None
        return MailItem.from_outlook_item(items.Item(index + 1))

    def get_subfolder(self, folder_name: str) -> Folder | None:
        """Get subfolder by name."""
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        normalized_folder_name: str = folder_name.lower()
        for folder in self.subfolders:
            if (folder.name or "").lower() == normalized_folder_name:
                return folder
        return None

    def create_subfolder(self, folder_name: str) -> Folder | None:
        """Create a new subfolder with the given name."""
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        if not folder_name.strip():
            raise ValueError("folder_name cannot be empty")
        folders = self._ol_folder_item.Folders
        if folders is None:
            raise AttributeError("Folder does not expose a Folders collection")
        return Folder(folders.Add(folder_name))

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
        return mail_item.move(destination)

    def delete_item(self, mail_item: MailItem) -> None:
        """Delete a MailItem from the folder."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def delete(self) -> None:
        """Delete this folder."""
        self._ol_folder_item.Delete()

    def __iter__(self):
        yield from self.list_messages()

    def __repr__(self) -> str:
        """Return repr string for the folder."""
        name = self.name
        folder_path = self.folder_path
        return f"Folder({name=!r}, {folder_path=!r})"

    def __str__(self) -> str:
        """Return string representation of the folder."""
        return self.__repr__()
