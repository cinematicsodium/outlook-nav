import logging
from collections.abc import Iterator

from constants import FolderEnum

from models.folder import Folder
from protocols import NamespaceProtocol

logger = logging.getLogger(__name__)


class DefaultFolders:
    FOLDER_NAME_MAP: dict[str, FolderEnum] = {
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
        "managed_email": FolderEnum.MANAGED_EMAIL,
        "drafts": FolderEnum.DRAFTS,
    }

    def __init__(self, namespace: NamespaceProtocol):
        self.namespace = namespace
        self._cache: dict[FolderEnum, Folder | None] = {}

    @property
    def deleted_items(self) -> Folder | None:
        return self.resolve(FolderEnum.DELETED_ITEMS)

    @property
    def outbox(self) -> Folder | None:
        return self.resolve(FolderEnum.OUTBOX)

    @property
    def sent_mail(self) -> Folder | None:
        return self.resolve(FolderEnum.SENT_MAIL)

    @property
    def inbox(self) -> Folder | None:
        return self.resolve(FolderEnum.INBOX)

    @property
    def calendar(self) -> Folder | None:
        return self.resolve(FolderEnum.CALENDAR)

    @property
    def contacts(self) -> Folder | None:
        return self.resolve(FolderEnum.CONTACTS)

    @property
    def journal(self) -> Folder | None:
        return self.resolve(FolderEnum.JOURNAL)

    @property
    def notes(self) -> Folder | None:
        return self.resolve(FolderEnum.NOTES)

    @property
    def conflicts(self) -> Folder | None:
        return self.resolve(FolderEnum.CONFLICTS)

    @property
    def local_failures(self) -> Folder | None:
        return self.resolve(FolderEnum.LOCAL_FAILURES)

    @property
    def junk(self) -> Folder | None:
        return self.resolve(FolderEnum.JUNK)

    @property
    def managed_email(self) -> Folder | None:
        return self.resolve(FolderEnum.MANAGED_EMAIL)

    @property
    def drafts(self) -> Folder | None:
        return self.resolve(FolderEnum.DRAFTS)

    def resolve(self, target: FolderEnum, *, force_refresh: bool = False) -> Folder | None:
        if not force_refresh and target in self._cache:
            return self._cache[target]
        try:
            folder = self.namespace.GetDefaultFolder(target.value)
            resolved = Folder(folder, target)
            self._cache[target] = resolved
            return resolved
        except Exception:
            logger.exception("Error resolving default folder '%s'", target)
            self._cache[target] = None
            return None

    @property
    def all(self) -> list[Folder | None]:
        return [self.resolve(folder_enum) for folder_enum in self.FOLDER_NAME_MAP.values()]

    def to_dict(self) -> dict[str, Folder | None]:
        return {
            name: self.resolve(folder_enum) for name, folder_enum in self.FOLDER_NAME_MAP.items()
        }

    def as_dict(self) -> dict[str, Folder | None]:
        return self.to_dict()

    def __iter__(self) -> Iterator[Folder | None]:
        yield from self.all

    def clear_cache(self) -> None:
        self._cache.clear()
