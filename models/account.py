from __future__ import annotations

from collections.abc import Iterable

from outlook.enums import OutlookItemClass
from outlook.models.folder import Folder
from outlook.protocols import OlAccount, OlFolder, OlStore
from outlook.utils import is_valid_ol_item


class Account:
    def __init__(self, ol_acct_item: OlAccount):
        self.ol_account_item = ol_acct_item

    @classmethod
    def from_outlook_item(cls, ol_acct_item: OlAccount):
        if not is_valid_ol_item(ol_acct_item, OutlookItemClass.ACCOUNT):
            return None
        return cls(ol_acct_item)

    @property
    def name(self) -> str:
        return self.ol_account_item.DisplayName

    @property
    def store(self) -> OlStore:
        return self.ol_account_item.DeliveryStore

    @property
    def root_folder(self) -> OlFolder:
        return self.store.GetRootFolder()

    @property
    def folders(self):
        try:
            ol_folders = self.root_folder.Folders
            if not (count := ol_folders.Count):
                return []
            ol_accts = [ol_folders.Item(i) for i in range(1, count + 1)]
            folders = [Folder.from_outlook_item(f) for f in ol_accts]
            return [f for f in folders if f]
        except Exception:
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

    def __str__(self):
        return self.name

    def __repr__(self):
        return f"Account({self.name})"
