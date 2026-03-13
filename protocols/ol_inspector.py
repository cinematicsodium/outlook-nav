from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_editor import OlEditor


class OlInspector(OlItem):
    WordEditor: OlEditor
