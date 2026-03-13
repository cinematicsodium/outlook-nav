from __future__ import annotations

from collections.abc import Iterator

from ..enums import FolderType, ItemType
from ..logger import log
from ..models.folder import Folder
from ..protocols import OlNamespace
from ..utils import is_accessible_ol_item


class DefaultFolders:
    """Manages Outlook default folders."""

    def __init__(self, namespace: OlNamespace):
        """Initialize with Outlook namespace."""
        self._namespace = namespace
        self._cache: dict[FolderType, Folder | None] = {}

        if not self.is_accessible_namespace(namespace):
            raise ValueError("Provided namespace is not accessible or valid.")

    @classmethod
    def from_namespace(cls, namespace: OlNamespace) -> DefaultFolders:
        if not cls.is_accessible_namespace(namespace):
            return None
        return cls(namespace)

    @classmethod
    def is_accessible_namespace(cls, item: OlNamespace) -> bool:
        return is_accessible_ol_item(
            item=item,
            target_type=ItemType.NAMESPACE,
        )

    @staticmethod
    def folder_name_map() -> dict[str, FolderType]:
        """Return mapping of folder names to enums."""
        return {
            "deleted_items": FolderType.DELETED_ITEMS,
            "outbox": FolderType.OUTBOX,
            "sent_mail": FolderType.SENT_MAIL,
            "inbox": FolderType.INBOX,
            "calendar": FolderType.CALENDAR,
            "contacts": FolderType.CONTACTS,
            "journal": FolderType.JOURNAL,
            "notes": FolderType.NOTES,
            "conflicts": FolderType.CONFLICTS,
            "local_failures": FolderType.LOCAL_FAILURES,
            "junk": FolderType.JUNK,
            "drafts": FolderType.DRAFTS,
        }

    @property
    def deleted_items(self) -> Folder | None:
        """Return Deleted Items folder."""
        return self._get_default_folder("deleted_items")

    @property
    def outbox(self) -> Folder | None:
        """Return Outbox folder."""
        return self._get_default_folder("outbox")

    @property
    def sent_mail(self) -> Folder | None:
        """Return Sent Mail folder."""
        return self._get_default_folder("sent_mail")

    @property
    def inbox(self) -> Folder | None:
        """Return Inbox folder."""
        return self._get_default_folder("inbox")

    @property
    def calendar(self) -> Folder | None:
        """Return Calendar folder."""
        return self._get_default_folder("calendar")

    @property
    def contacts(self) -> Folder | None:
        """Return Contacts folder."""
        return self._get_default_folder("contacts")

    @property
    def journal(self) -> Folder | None:
        """Return Journal folder."""
        return self._get_default_folder("journal")

    @property
    def notes(self) -> Folder | None:
        """Return Notes folder."""
        return self._get_default_folder("notes")

    @property
    def conflicts(self) -> Folder | None:
        """Return Conflicts folder."""
        return self._get_default_folder("conflicts")

    @property
    def local_failures(self) -> Folder | None:
        """Return Local Failures folder."""
        return self._get_default_folder("local_failures")

    @property
    def junk(self) -> Folder | None:
        """Return Junk folder."""
        return self._get_default_folder("junk")

    @property
    def drafts(self) -> Folder | None:
        """Return Drafts folder."""
        return self._get_default_folder("drafts")

    @property
    def all(self) -> list[Folder]:
        """Return all default folders as a list."""
        return list(self._to_dict().values())

    def _to_dict(self) -> dict[str, Folder]:
        """Return dictionary of default folders."""
        data: dict[str, Folder | None] = {
            name: self._get_default_folder(name) for name in self.folder_name_map()
        }
        return {k: v for k, v in data.items() if v is not None}

    def as_dict(self) -> dict[str, Folder]:
        """Return default folders as a dictionary."""
        return self._to_dict()

    def _to_list(self) -> list[Folder]:
        """Return list of default folders."""
        return list(self._to_dict().values())

    def as_list(self) -> list[Folder]:
        """Return default folders as a list."""
        return self._to_list()

    def _get_default_folder(self, folder_name: str) -> Folder | None:
        """Return default folder by name."""
        mapping = self.folder_name_map()

        if not folder_name or folder_name not in mapping:
            return None

        if folder_name in self._cache:
            return self._cache[folder_name]

        target = mapping[folder_name]
        try:
            ol_folder = self._namespace.GetDefaultFolder(target)
            folder = Folder(ol_folder)
            self._cache[target] = folder
            return folder
        except Exception:
            log.error("Error resolving default folder '%s'", target)
            self._cache[target] = None
            return None

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

    def _clear_cache(self) -> None:
        """Clear folder cache."""
        self._cache.clear()
