import logging
from collections.abc import Iterator

from outlook.constants import DefaultFolderEnum
from outlook.models.folder import Folder
from outlook.protocols import OlNamespace

logger = logging.getLogger(__name__)


class DefaultFolders:
    FOLDER_NAME_MAP: dict[str, DefaultFolderEnum] = {
        "deleted_items": DefaultFolderEnum.DELETED_ITEMS,
        "outbox": DefaultFolderEnum.OUTBOX,
        "sent_mail": DefaultFolderEnum.SENT_MAIL,
        "inbox": DefaultFolderEnum.INBOX,
        "calendar": DefaultFolderEnum.CALENDAR,
        "contacts": DefaultFolderEnum.CONTACTS,
        "journal": DefaultFolderEnum.JOURNAL,
        "notes": DefaultFolderEnum.NOTES,
        "conflicts": DefaultFolderEnum.CONFLICTS,
        "local_failures": DefaultFolderEnum.LOCAL_FAILURES,
        "junk": DefaultFolderEnum.JUNK,
        "drafts": DefaultFolderEnum.DRAFTS,
    }

    def __init__(self, namespace: OlNamespace):
        self.namespace = namespace
        self._cache: dict[DefaultFolderEnum, Folder | None] = {}

    @property
    def deleted_items(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.DELETED_ITEMS)

    @property
    def outbox(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.OUTBOX)

    @property
    def sent_mail(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.SENT_MAIL)

    @property
    def inbox(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.INBOX)

    @property
    def calendar(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.CALENDAR)

    @property
    def contacts(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.CONTACTS)

    @property
    def journal(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.JOURNAL)

    @property
    def notes(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.NOTES)

    @property
    def conflicts(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.CONFLICTS)

    @property
    def local_failures(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.LOCAL_FAILURES)

    @property
    def junk(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.JUNK)

    @property
    def drafts(self) -> Folder | None:
        return self.resolve(DefaultFolderEnum.DRAFTS)

    def resolve(
        self, target: DefaultFolderEnum, *, force_refresh: bool = False
    ) -> Folder | None:
        if not force_refresh and target in self._cache:
            return self._cache[target]
        try:
            folder = self.namespace.GetDefaultFolder(target.value)
            resolved = Folder(folder)
            self._cache[target] = resolved
            return resolved
        except Exception:
            logger.error("Error resolving default folder '%s'", target)
            self._cache[target] = None
            return None

    @property
    def all(self) -> list[Folder]:
        return list(self.to_dict().values())

    def to_dict(self) -> dict[str, Folder]:
        data = {
            name: self.resolve(folder_enum)
            for name, folder_enum in self.FOLDER_NAME_MAP.items()
        }
        return {k: v for k, v in data.items() if v is not None}

    def as_dict(self) -> dict[str, Folder]:
        return self.to_dict()

    def __iter__(self) -> Iterator[Folder]:
        yield from self.all

    def __str__(self):
        items = "; ".join(str(v) for v in self.as_dict().values() if v)
        return items

    def _clear_cache(self) -> None:
        self._cache.clear()
