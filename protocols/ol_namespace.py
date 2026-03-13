from __future__ import annotations

from typing import TYPE_CHECKING, Any

from ..types import T
from .ol_item import OlItem

if TYPE_CHECKING:
    from ..enums import FolderType
    from .ol_account import OlAccount
    from .ol_collection import OlCollection
    from .ol_folder import OlFolder
    from .ol_recipient import OlRecipient


class OlNamespace(OlItem):
    Accounts: OlCollection[OlAccount]
    Folders: OlCollection[OlFolder]

    def GetDefaultFolder(self, folder_type: FolderType | int) -> OlFolder: ...
    def CreateRecipient(self, recipient_name: str) -> OlRecipient: ...
    def GetItemFromID(self, entry_id: str, store_id: Any = None) -> T: ...
