from __future__ import annotations

import logging
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any

from outlook.constants import DefaultFolderEnum, OutlookItemClass
from outlook.models.mail_item import MailItem
from outlook.protocols import OlFolder
from outlook.utils import is_valid_ol_item

logger = logging.getLogger(__name__)


@dataclass
class Folder:
    def __init__(self, outlook_item: OlFolder):
        self._ol_folder_item = outlook_item

    @classmethod
    def from_outlook_item(cls, item: Any) -> Folder | None:
        if not cls.is_valid_folder_item(item):
            return None

        return cls(item)

    @staticmethod
    def interface_properties():
        return ("Name", "Items", "Folders")

    @staticmethod
    def is_valid_folder_item(item: Any):
        return (
            is_valid_ol_item(
                item=item,
                target_type=OutlookItemClass.FOLDER,
                properties=Folder.interface_properties(),
            ),
        )

    @property
    def name(self) -> str:
        return self._ol_folder_item.Name

    @property
    def mail_items(self) -> list[MailItem]:

        folder_items = self._ol_folder_item.Items
        item_count = folder_items.Count

        if not (folder_items and item_count):
            return []

        mail_items = []
        outlook_items = [folder_items.Item(idx) for idx in range(1, item_count + 1)]
        mail_items = [MailItem.from_outlook_item(item) for item in outlook_items]
        valid_items = [item for item in mail_items if item]
        return valid_items

    @property
    def subfolders(self) -> list[Folder]:
        try:
            folders = self._ol_folder_item.Folders
            count = folders.Count

            if not (folders and count):
                return []

            outlook_items = [folders.Item(idx) for idx in range(1, count + 1)]

            folder_items = [
                Folder.from_outlook_item(folder)
                for folder in outlook_items
                if folder.Class == OutlookItemClass.FOLDER
            ]

            return folder_items

        except Exception:
            return []

    def get_item(self, index: int) -> MailItem | None:
        if index < 0:
            raise ValueError("index must be >= 0")

        items = self.mail_items
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
        return mail_item.move_to(destination)

    def delete_item(self, mail_item: MailItem) -> None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def extend_items(self, values: Iterable[MailItem]) -> list[MailItem]:
        return list(values)

    def delete(self) -> None:
        if self._is_default_folder(self._ol_folder_item):
            raise ValueError(f"Cannot delete default folder: {self.folder_type}")
        try:
            self._ol_folder_item.Delete()
        except Exception:
            logger.error("Error deleting folder '%s'", self.name)

    @staticmethod
    def _is_default_folder(item: OlFolder):
        try:
            return bool(DefaultFolderEnum(item.Class))
        except Exception:
            return False

    def __str__(self):
        return self.name

    def __repr__(self):
        return f"Folder(name={self.name})"
