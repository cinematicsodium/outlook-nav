from typing import Any, Protocol, TypeAlias

from .types import T


class OlCollection(Protocol[T]):
    """Structural type for a one-indexed Outlook COM collection."""

    Count: int

    def Item(self, index: int, /) -> T: ...


OlAccount: TypeAlias = Any
OlAddressEntry: TypeAlias = Any
OlApplication: TypeAlias = Any
OlFolder: TypeAlias = Any
OlMailItem: TypeAlias = Any
OlNamespace: TypeAlias = Any
OlObject: TypeAlias = Any
OlStore: TypeAlias = Any
