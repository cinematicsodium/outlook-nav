from __future__ import annotations

from collections.abc import Iterator
from typing import Any, Protocol

from ..types import T
from .ol_item import OlItem


class OlCollection(OlItem, Protocol[T]):
    """Generic protocol for COM collections."""

    Count: int

    def Item(self, index: int | str, /) -> T: ...
    def ResolveAll(self, /) -> T: ...
    def Add(self, obj: Any, /) -> T: ...
    def __iter__(self) -> Iterator[T]: ...
