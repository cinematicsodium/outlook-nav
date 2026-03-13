from __future__ import annotations

from .ol_item import OlItem


class OlAttachment(OlItem):
    DisplayName: str
    FileName: str
    PathName: str
    Position: int
    Size: int
    Type: int

    def SaveAsFile(self, path: str) -> None: ...
    def Delete(self) -> None: ...
