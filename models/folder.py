from __future__ import annotations

import logging
from collections.abc import Iterable
from typing import cast

from constants import FolderEnum
from models.mail_item import MailItem
from protocols import FolderProtocol
from utils import resolve_method, resolve_property

logger = logging.getLogger(__name__)


class Folder:
    def __init__(self, folder_item: FolderProtocol, folder_type: FolderEnum | None = None):
        self.folder = folder_item
        self.type = folder_type

    @property
    def name(self) -> str | None:
        return cast(str | None, resolve_property(self.folder, "Name"))

    @property
    def items(self) -> list[MailItem]:
        items = resolve_property(self.folder, "Items")
        if items is None:
            return []

        # Outlook COM collections can be iterable, but some only expose Count/Item.
        try:
            return [MailItem(item) for item in list(items)]
        except Exception:
            count = resolve_property(items, "Count")
            if not isinstance(count, int) or count <= 0:
                return []

            wrapped_items: list[MailItem] = []
            for index in range(1, count + 1):
                try:
                    wrapped_items.append(MailItem(resolve_method(items, "Item", index)))
                except Exception:
                    logger.exception(
                        "Error wrapping folder item at COM index %s from '%s'",
                        index,
                        self.name,
                    )
            return wrapped_items

    @property
    def subfolders(self) -> list[Folder]:
        folders = resolve_property(self.folder, "Folders")
        if folders is None:
            return []
        return [Folder(folder) for folder in list(folders)]

    def get_item(self, index: int) -> MailItem | None:
        if index < 0:
            raise ValueError("index must be >= 0")

        items = self.items
        if index >= len(items):
            return None
        return items[index]

    def get_subfolder(self, folder_name: str) -> Folder | None:
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")

        target = folder_name.lower()
        for folder in self.subfolders:
            if (folder.name or "").lower() == target:
                return folder
        return None

    def create_subfolder(self, folder_name: str) -> Folder | None:
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        if not folder_name.strip():
            raise ValueError("folder_name cannot be empty")
        try:
            folders = resolve_property(self.folder, "Folders")
            if folders is None:
                raise AttributeError("Folder does not expose a Folders collection")
            created = resolve_method(folders, "Add", folder_name)
            return Folder(created)
        except Exception:
            logger.exception("Error creating subfolder '%s' in '%s'", folder_name, self.name)
            return None

    def delete_subfolder(self, folder_name: str) -> bool:
        folder = self.get_subfolder(folder_name)
        if folder is None:
            return False
        folder.delete()
        return True

    def move_item(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        if not isinstance(destination, Folder):
            raise ValueError("destination must be a Folder")
        return mail_item.move(destination)

    def delete_item(self, mail_item: MailItem) -> None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def extend_items(self, values: Iterable[MailItem]) -> list[MailItem]:
        return list(values)

    def delete(self) -> None:
        if self.type is not None:
            raise ValueError(f"Cannot delete default folder: {self.type}")
        try:
            resolve_method(self.folder, "Delete")
        except Exception:
            logger.exception("Error deleting folder '%s'", self.name)
