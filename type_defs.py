from typing import TYPE_CHECKING, TypeAlias, TypeVar

if TYPE_CHECKING:
    from .models.node import ItemModel

LowerStr: TypeAlias = str

T = TypeVar("T", covariant=True)

ModelT = TypeVar("ModelT", bound="ItemModel")

RawT = TypeVar("RawT")
