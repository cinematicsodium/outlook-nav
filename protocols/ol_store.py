from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_folder import OlFolder


class OlStore(OlItem):
    def GetRootFolder(self) -> OlFolder: ...
