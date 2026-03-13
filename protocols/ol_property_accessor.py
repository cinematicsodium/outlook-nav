from __future__ import annotations

from typing import Any

from .ol_item import OlItem


class OlPropertyAccessor(OlItem):
    def GetProperty(self, schema_name: str) -> Any: ...
    def SetProperty(self, schema_name: str, value: Any) -> None: ...
