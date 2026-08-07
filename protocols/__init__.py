from typing import Any, Protocol, TypeAlias, TypeVar

T = TypeVar("T", covariant=True)


class OlCollection(Protocol[T]):
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

__all__ = [
    "OlAccount",
    "OlAddressEntry",
    "OlApplication",
    "OlCollection",
    "OlFolder",
    "OlMailItem",
    "OlNamespace",
    "OlObject",
    "OlStore",
]
