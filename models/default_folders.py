from __future__ import annotations

import logging
from collections.abc import Iterator
from functools import cached_property

from ..enums import FolderEnum, ItemType
from ..models.folder import Folder
from ..protocols import OlNamespace
from .base import ItemModel

log = logging.getLogger(__name__)


class DefaultFolders(ItemModel):
    """Manages Outlook default folders."""

    item_type = ItemType.NAMESPACE
    inaccessible_error_message = "Provided Outlook item is not an accessible namespace."
    _NAME_MAP: dict[str, FolderEnum] = {
        "deleted_items": FolderEnum.DELETED_ITEMS,
        "outbox": FolderEnum.OUTBOX,
        "sent_mail": FolderEnum.SENT_MAIL,
        "inbox": FolderEnum.INBOX,
        "calendar": FolderEnum.CALENDAR,
        "contacts": FolderEnum.CONTACTS,
        "journal": FolderEnum.JOURNAL,
        "notes": FolderEnum.NOTES,
        "conflicts": FolderEnum.CONFLICTS,
        "local_failures": FolderEnum.LOCAL_FAILURES,
        "junk": FolderEnum.JUNK,
        "drafts": FolderEnum.DRAFTS,
    }

    def __init__(self, namespace: OlNamespace):
        """Initialize with Outlook namespace."""
        super().__init__(namespace)
        self._ol_namespace = namespace
        self._cache: dict[FolderEnum, Folder] = {}

    @cached_property
    def deleted_items(self) -> Folder | None:
        """Return Deleted Items folder."""
        return self.get("deleted_items")

    @cached_property
    def outbox(self) -> Folder | None:
        """Return Outbox folder."""
        return self.get("outbox")

    @cached_property
    def sent_mail(self) -> Folder | None:
        """Return Sent Mail folder."""
        return self.get("sent_mail")

    @cached_property
    def inbox(self) -> Folder | None:
        """Return Inbox folder."""
        return self.get("inbox")

    @cached_property
    def calendar(self) -> Folder | None:
        """Return Calendar folder."""
        return self.get("calendar")

    @cached_property
    def contacts(self) -> Folder | None:
        """Return Contacts folder."""
        return self.get("contacts")

    @cached_property
    def journal(self) -> Folder | None:
        """Return Journal folder."""
        return self.get("journal")

    @cached_property
    def notes(self) -> Folder | None:
        """Return Notes folder."""
        return self.get("notes")

    @cached_property
    def conflicts(self) -> Folder | None:
        """Return Conflicts folder."""
        return self.get("conflicts")

    @cached_property
    def local_failures(self) -> Folder | None:
        """Return Local Failures folder."""
        return self.get("local_failures")

    @cached_property
    def junk(self) -> Folder | None:
        """Return Junk folder."""
        return self.get("junk")

    @cached_property
    def drafts(self) -> Folder | None:
        """Return Drafts folder."""
        return self.get("drafts")

    @property
    def all(self) -> list[Folder]:
        """Return all default folders as a list."""
        return self.as_list()

    def as_dict(self) -> dict[str, Folder]:
        """Return default folders as a dictionary."""
        return {
            name: resolved
            for name in self._NAME_MAP
            if (resolved := self.get(name)) is not None
        }

    def as_list(self) -> list[Folder]:
        """Return default folders as a list."""
        return list(self.as_dict().values())

    def get(self, folder: str | FolderEnum) -> Folder | None:
        """Return a default folder by name or enum."""
        folder_type = self._resolve_folder_type(folder)
        if folder_type is None:
            if folder is not None:
                log.warning("Unknown default folder identifier: %r", folder)
            return None
        if folder_type in self._cache:
            return self._cache[folder_type]
        ol_folder = self._ol_namespace.GetDefaultFolder(folder_type)
        resolved_folder = Folder(ol_folder)
        self._cache[folder_type] = resolved_folder
        return resolved_folder

    def _resolve_folder_type(self, folder: str | FolderEnum) -> FolderEnum | None:
        if isinstance(folder, FolderEnum):
            return folder
        if not folder:
            return None
        return self._NAME_MAP.get(folder.strip().lower())

    def __iter__(self) -> Iterator[Folder]:
        """Iterate over default folders."""
        yield from self.all

    def __str__(self) -> str:
        """Return string representation."""
        items = ", ".join(str(f) for f in self.as_list())
        return items

    def __repr__(self) -> str:
        """Return repr representation."""
        folders = ", ".join(repr(f) for f in self.as_list())
        return f"DefaultFolders({folders})"
