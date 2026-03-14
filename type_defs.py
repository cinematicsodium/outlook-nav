from typing import TypeAlias, TypeVar

from .models.node import ItemModel

LowerStr: TypeAlias = str

T = TypeVar("T", covariant=True)

ModelT = TypeVar("ModelT", bound=ItemModel)
RawT = TypeVar("RawT")
