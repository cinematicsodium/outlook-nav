from __future__ import annotations

from collections.abc import Iterator

from ..enums import FolderType, ItemType
from ..logger import log
from ..models.folder import Folder
from ..protocols import OlNamespace
from ..utils import is_accessible_ol_item


class DefaultFolders:
    """Manages Outlook default folders."""

    _NAME_MAP: dict[str, FolderType] = {
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

    def __init__(self, namespace: OlNamespace):
        """Initialize with Outlook namespace."""
        self._namespace = namespace
        self._cache: dict[FolderType, Folder | None] = {}

        if not self.is_accessible_namespace(namespace):
            raise ValueError("Provided namespace is not accessible or valid.")

    @classmethod
    def from_namespace(cls, namespace: OlNamespace) -> DefaultFolders | None:
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
        return dict(DefaultFolders._NAME_MAP)

    @property
    def deleted_items(self) -> Folder | None:
        """Return Deleted Items folder."""
        return self.get("deleted_items")

    @property
    def outbox(self) -> Folder | None:
        """Return Outbox folder."""
        return self.get("outbox")

    @property
    def sent_mail(self) -> Folder | None:
        """Return Sent Mail folder."""
        return self.get("sent_mail")

    @property
    def inbox(self) -> Folder | None:
        """Return Inbox folder."""
        return self.get("inbox")

    @property
    def calendar(self) -> Folder | None:
        """Return Calendar folder."""
        return self.get("calendar")

    @property
    def contacts(self) -> Folder | None:
        """Return Contacts folder."""
        return self.get("contacts")

    @property
    def journal(self) -> Folder | None:
        """Return Journal folder."""
        return self.get("journal")

    @property
    def notes(self) -> Folder | None:
        """Return Notes folder."""
        return self.get("notes")

    @property
    def conflicts(self) -> Folder | None:
        """Return Conflicts folder."""
        return self.get("conflicts")

    @property
    def local_failures(self) -> Folder | None:
        """Return Local Failures folder."""
        return self.get("local_failures")

    @property
    def junk(self) -> Folder | None:
        """Return Junk folder."""
        return self.get("junk")

    @property
    def drafts(self) -> Folder | None:
        """Return Drafts folder."""
        return self.get("drafts")

    @property
    def all(self) -> list[Folder]:
        """Return all default folders as a list."""
        return list(self._to_dict().values())

    def _to_dict(self) -> dict[str, Folder]:
        """Return dictionary of default folders."""
        data: dict[str, Folder | None] = {
            name: self.get(name) for name in self.folder_name_map()
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

    def get(self, folder: str | FolderType) -> Folder | None:
        """Return a default folder by name or enum."""
        folder_type = self._resolve_folder_type(folder)
        if folder_type is None:
            return None

        if folder_type in self._cache:
            return self._cache[folder_type]

        try:
            ol_folder = self._namespace.GetDefaultFolder(folder_type)
            resolved_folder = Folder(ol_folder)
            self._cache[folder_type] = resolved_folder
            return resolved_folder
        except Exception:
            log.error("Error resolving default folder '%s'", folder_type)
            self._cache[folder_type] = None
            return None

    def _get_default_folder(self, folder: str | FolderType) -> Folder | None:
        """Backward-compatible wrapper for retrieving a default folder."""
        return self.get(folder)

    def _resolve_folder_type(self, folder: str | FolderType) -> FolderType | None:
        if isinstance(folder, FolderType):
            return folder
        if not folder:
            return None
        return self.folder_name_map().get(folder)

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
