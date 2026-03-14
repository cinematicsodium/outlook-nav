from __future__ import annotations

from collections.abc import Iterable
from enum import IntEnum
from typing import TYPE_CHECKING, Any, TypeVar, cast, overload

if TYPE_CHECKING:
    from .models.node import ItemModel

try:
    import pywintypes  # type: ignore
except ImportError:  # pragma: no cover - depends on Windows/pywin32
    pywintypes = None

from .constants import SMTP_ADDRESS_SCHEMA, UNSET
from .enums import AddressUserType
from .protocols import OlAddressEntry, OlCollection, OlItem
from .type_defs import LowerStr, ModelT, RawT, T


def resolve_property(obj: Any, property_name: str, new_value: Any = UNSET) -> T | None:
    """Resolves a property on an object, optionally setting it to a new value."""
    try:
        if not hasattr(obj, property_name):
            if new_value is UNSET:
                return None
            raise AttributeError(f"Object does not have property '{property_name}'.")

        if new_value is not UNSET:
            try:
                setattr(obj, property_name, new_value)
            except Exception:
                pass

        if not hasattr(obj, property_name):
            return None

        return cast(T, getattr(obj, property_name))
    except Exception as e:
        if is_interface_error(e):
            return None
        raise


def resolve_method(obj: Any, method_name: str, *args: Any, **kwargs: Any) -> T | None:
    """Resolves a method on an object and calls it with the given arguments."""
    try:
        if not hasattr(obj, method_name):
            raise AttributeError(f"Object does not have method '{method_name}'.")

        method = getattr(obj, method_name)
        if not callable(method):
            raise AttributeError(f"'{method_name}' is not a callable method.")

        return cast(T, method(*args, **kwargs))
    except Exception as e:
        if is_interface_error(e):
            return None
        raise


def is_interface_error(e: Exception) -> bool:
    if pywintypes is None:
        return False
    return isinstance(e, pywintypes.com_error)


def is_accessible_ol_item(
    item: OlItem, target_type: IntEnum, properties: Iterable[str] | None = None
) -> bool:
    if not hasattr(item, "Class"):
        return False
    rules = (
        item.Class == target_type,
        is_ol_item_accessible(item, properties),
    )
    return all(rules)


def is_ol_item_accessible(
    item: OlItem, properties: Iterable[str] | None = None
) -> bool:
    """
    Attempts to access standard attributes to ensure the item isn't
    restricted by security systems.
    """
    if not properties:
        return True
    try:
        for property in properties:
            getattr(item, property)
        return True
    except Exception:
        return False


@overload
def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: None = None,
) -> list[T]: ...


@overload
def unpack_collection(
    collection: OlCollection[RawT] | None,
    *,
    transformer: type[ModelT],
) -> list[ModelT]: ...


def unpack_collection(
    collection: OlCollection[T] | None,
    *,
    transformer: type[ItemModel] | None = None,
) -> list[T] | list[ItemModel]:
    if collection is None:
        return []

    count = collection.Count
    if count == 0:
        return []

    items = [collection.Item(i + 1) for i in range(count)]
    if transformer is None:
        return cast(list[T], [i for i in items if i is not None])

    transformed_items = [transformer.from_outlook_item(item) for item in items]
    return cast(list[ModelT], [i for i in transformed_items if i is not None])


def is_exchange_user(user: OlAddressEntry) -> bool:
    rules = [
        user.Type == "EX",
        user.AddressEntryUserType == AddressUserType.EXCHANGE_REMOTE_USER,
        user.AddressEntryUserType == AddressUserType.EXCHANGE_USER,
    ]
    return any(rules)


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

        exch_user = user.GetExchangeUser()
        if exch_user:
            return str(exch_user.PrimarySmtpAddress).lower()

        try:
            address = user.PropertyAccessor.GetProperty(SMTP_ADDRESS_SCHEMA)
            return str(address).lower()
        except Exception:
            return str(user.Address).lower()
    except Exception:
        return ""
