from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_collection import OlCollection


class OlEditor(OlItem):
    Tables: OlCollection

    def Range(self): ...
