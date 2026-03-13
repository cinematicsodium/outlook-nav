from __future__ import annotations

from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from .ol_application import OlApplication


class OlDispatch(Protocol):
    def Dispatch(self, app_name: str) -> OlApplication: ...
