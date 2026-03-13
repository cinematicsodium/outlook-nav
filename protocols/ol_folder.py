from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_collection import OlCollection
    from .ol_mail_item import OlMailItem


class OlFolder(OlItem):
    Name: str
    EntryID: str
    Folders: OlCollection[OlFolder]
    Items: OlCollection[OlMailItem]

    def Delete(self) -> None: ...
    def CopyTo(self, destination: OlFolder) -> OlFolder: ...
