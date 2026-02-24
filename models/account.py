from __future__ import annotations

from collections.abc import Iterable
from typing import Any

from models.folder import Folder


class MailboxAccount:
    def __init__(self, mailbox: Any):
        self.mailbox = mailbox

    @property
    def name(self) -> str:
        return str(getattr(self.mailbox, "Name", ""))

    @property
    def folders(self) -> list[Folder]:
        folder_collection = getattr(self.mailbox, "Folders", None)
        if folder_collection is None:
            return []
        try:
            return [Folder(folder) for folder in list(folder_collection)]
        except TypeError:
            return []

    def get_folder(self, folder_name: str) -> Folder | None:
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")
        target = folder_name.lower()
        for folder in self.folders:
            if (folder.name or "").lower() == target:
                return folder
        return None

    def find_folder(self, path: str | Iterable[str]) -> Folder | None:
        if isinstance(path, str):
            parts = [segment.strip() for segment in path.split("/") if segment.strip()]
        else:
            parts = [str(segment).strip() for segment in path if str(segment).strip()]

        if not parts:
            return None

        current = self.get_folder(parts[0])
        for segment in parts[1:]:
            if current is None:
                return None
            current = current.get_subfolder(segment)
        return current
