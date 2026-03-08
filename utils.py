import re
from collections.abc import Iterable
from datetime import datetime
from enum import IntEnum
from pathlib import Path
from typing import Any, Callable, TypeAlias, cast

import pywintypes  # type: ignore

from outlook.constants import (
    DIGIT_REGEX,
    EMAIL_REGEX,
    SMTP_ADDRESS_SCHEMA,
    TIMEZONE_OFFSET,
)
from outlook.enums import AddressUserEnum
from outlook.protocols import OlAddressEntry, OlItem

LowerStr: TypeAlias = str
_UNSET = object()


def ensure_list(obj: Any, key: Callable | None = None) -> list[Any]:
    items = list(obj) if isinstance(obj, (list, set, tuple)) else [obj]
    if key is None:
        return items
    return [key(i) for i in items]


def _parse_email_values(email_values: list[str], pattern=re.compile(r"\s*[;,]\s*")):
    parsed_emails: list[str] = []

    for email in email_values:
        if not email:
            continue
        parts = pattern.split(email)
        parsed_emails.extend(p.strip() for p in parts if p.strip())

    return parsed_emails


def _validate_emails_impl(emails: list[str]):
    valid = []
    errors = []

    for email in emails:
        search = EMAIL_REGEX.search(email)
        if not search:
            err = f"Invalid email address: {email} (format check failed)"
            errors.append(err)
            continue
        valid.append(search.group())

    if errors:
        err_msg = "Email validation errors:\n" + "\n".join(errors)
        raise ValueError(err_msg)
    return "; ".join(valid)


def validate_email(emails: str | Iterable[str] | None) -> LowerStr:
    """Validate one or more emails and return Outlook-style recipient string."""
    if not emails:
        return None
    email_values: list[str] = ensure_list(emails, key=lambda x: str(x).lower().strip())
    parsed_emails = _parse_email_values(email_values)
    valid_emails = _validate_emails_impl(parsed_emails)
    return valid_emails


def validate_paths(paths: str | Path | Iterable[str | Path]) -> list[Path]:
    """Validate one or more filesystem paths."""
    path_values = ensure_list(paths)

    valid = []
    errors = []

    for path in path_values:
        if not isinstance(path, (str, Path)):
            err = f"Invalid path: {path} (type: {type(path)})"
            errors.append(err)
            continue

        normalized_path = Path(path).expanduser().resolve()
        if not normalized_path.exists():
            err = f"Path does not exist: {normalized_path}"
            errors.append(err)
            continue

        valid.append(normalized_path)

    if errors:
        err_msg = "Path validation errors:\n" + "\n".join(errors)
        raise ValueError(err_msg)

    return valid


def apply_timezone_offset(dt: datetime) -> datetime:
    """Helper to consistently apply the global timezone offset."""
    return dt + TIMEZONE_OFFSET


def parse_date_string(date_str: str) -> datetime | None:
    """Handles the logic of converting strings to datetime objects."""
    try:
        iso_candidate = date_str.replace("Z", "+00:00")
        return datetime.fromisoformat(iso_candidate)
    except ValueError:
        pass

    return parse_date_value(date_str)


def validate_datetime(date_value: Any) -> datetime | None:
    """
    Validates and normalizes input into a datetime object with offset applied.
    Returns None if input is null or parsing fails.
    """
    if date_value is None:
        return None

    dt: datetime | None = (
        date_value
        if isinstance(date_value, datetime)
        else parse_date_string(str(date_value))
    )

    if dt:
        return apply_timezone_offset(dt)

    return None


def parse_date_value(input_str: str) -> datetime | None:
    digit_list = extract_digits(input_str)
    int_list = parse_digits(digit_list)
    while int_list:
        try:
            return datetime(*int_list)
        except Exception:
            int_list.pop()
    return None


def extract_digits(input_str: str) -> list[str]:
    return DIGIT_REGEX.findall(str(input_str))


def parse_digits(digit_list: list[str]) -> list[int]:
    try:
        return [int(d) for d in digit_list]
    except Exception:
        return []


def resolve_property(obj: Any, property_name: str, new_value: Any = _UNSET) -> Any:
    """Resolves a property on an object, optionally setting it to a new value."""
    try:
        if not hasattr(obj, property_name):
            if new_value is _UNSET:
                return None
            raise AttributeError(f"Object does not have property '{property_name}'.")

        if new_value is not _UNSET:
            try:
                setattr(obj, property_name, new_value)
            except Exception:
                pass

        if not hasattr(obj, property_name):
            return None

        return cast(Any, getattr(obj, property_name))
    except Exception as e:
        if is_interface_error(e):
            return None
        raise


def resolve_method(obj: Any, method_name: str, *args: Any, **kwargs: Any) -> Any:
    """Resolves a method on an object and calls it with the given arguments."""
    try:
        if not hasattr(obj, method_name):
            raise AttributeError(f"Object does not have method '{method_name}'.")

        method = getattr(obj, method_name)
        if not callable(method):
            raise AttributeError(f"'{method_name}' is not a callable method.")

        return method(*args, **kwargs)
    except Exception as e:
        if is_interface_error(e):
            return None
        raise


def is_exchange_user(user: OlAddressEntry):
    rules = [
        user.Type == "EX",
        user.AddressEntryUserType == AddressUserEnum.EXCHANGE_REMOTE_USER,
        user.AddressEntryUserType == AddressUserEnum.EXCHANGE_USER,
    ]
    return any(rules)


def get_smtp_address(user: OlAddressEntry):
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


def is_valid_ol_item(
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


def is_interface_error(e: Exception):
    return isinstance(e, pywintypes.com_error)
