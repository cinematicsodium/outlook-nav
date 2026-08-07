from __future__ import annotations
from collections.abc import Iterable
from enum import IntEnum
from typing import Callable, cast, overload

from .constants import _UNSET, SMTP_ADDRESS_SCHEMA
from .protocols import OlAddressEntry, OlCollection, OlObject
from .type_defs import LowerStr, ModelT, RawT, T


def is_accessible_ol_item(
    item: OlObject, target_type: IntEnum, properties: Iterable[str] | None = None
) -> bool:
    if item is None:
        return False
    item_type = getattr(item, "Class", _UNSET)
    if item_type is not _UNSET:
        return item_type == target_type
    properties = properties or ()
    try:
        for name in properties:
            getattr(item, name)
        return True
    except Exception:
        return False


@overload
def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: None = None,
    limit: int | None = None,
    predicate: Callable[[T], bool] | None = None,
) -> list[T]: ...
@overload
def unpack_collection(
    collection: OlCollection[RawT] | None,
    *,
    transformer: type[ModelT],
    limit: int | None = None,
    predicate: Callable[[RawT], bool] | None = None,
) -> list[ModelT]: ...
def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: type[ModelT] | None = None,
    limit: int | None = None,
    predicate: Callable[[T], bool] | None = None,
) -> list[T] | list[ModelT]:
    if collection is None or limit is not None and limit <= 0:
        return []
    result = []
    for idx in range(1, collection.Count + 1):
        item = collection.Item(idx)
        if item is None or predicate is not None and not predicate(item):
            continue
        value = item if transformer is None else transformer.from_outlook_item(item)
        if value is not None:
            result.append(value)
            if limit is not None and len(result) >= limit:
                break
    return cast(list[T] | list[ModelT], result)


def get_smtp_address(user: OlAddressEntry) -> LowerStr:
    """Returns the SMTP email address of the address entry.
    Args:
        user (OlAddressEntry): The address entry object.
    Returns:
        LowerStr: The SMTP email address.
    Notes:
        - For Exchange users, it attempts to get the primary SMTP address.
        - For non-Exchange users, it tries to retrieve the SMTP address via PropertyAccessor.
        - If all else fails, it falls back to the standard Address property.
    """
    try:
        if not user:
            return ""
        try:
            exch_user = user.GetExchangeUser()
        except Exception:
            exch_user = None
        if exch_user:
            return str(exch_user.PrimarySmtpAddress).lower()
        try:
            address = user.PropertyAccessor.GetProperty(SMTP_ADDRESS_SCHEMA)
            return str(address).lower()
        except Exception:
            return str(user.Address).lower()
    except Exception:
        return ""
