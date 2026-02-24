import re
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any, TypeAlias, cast

from constants import DIGIT_REGEX, EMAIL_REGEX

LowerStr: TypeAlias = str
_UNSET = object()


def ensure_list(obj: Any) -> list[Any]:
    """Return `obj` as a list, expanding non-string iterables."""
    if isinstance(obj, Iterable) and not isinstance(obj, (str, bytes, bytearray)):
        return list(obj)
    return [obj]


def validate_email(emails: str | Iterable[str]) -> LowerStr:
    """Validate one or more emails and return Outlook-style recipient string."""
    email_values = ensure_list(emails)

    valid = []
    errors = []

    for email in email_values:
        if not isinstance(email, str):
            err = f"Invalid email address: {email} (type: {type(email)})"
            errors.append(err)
            continue

        chunks = [email]
        if ";" in email or "," in email:
            chunks = [part.strip() for part in re.split(r"[;,]", email) if part.strip()]

        for chunk in chunks:
            normalized_email = chunk.strip().lower()
            if not EMAIL_REGEX.fullmatch(normalized_email):
                err = f"Invalid email address: {chunk} (format check failed)"
                errors.append(err)
                continue

            valid.append(normalized_email)

    if errors:
        err_msg = "Email validation errors:\n" + "\n".join(errors)
        raise ValueError(err_msg)

    return "; ".join(valid)


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


def validate_datetime(date_value: Any) -> datetime | Any:
    """Try to parse a datetime from COM/string values; return original on failure."""
    if isinstance(date_value, datetime):
        return date_value

    if isinstance(date_value, str):
        iso_candidate = date_value.replace("Z", "+00:00")
        try:
            return datetime.fromisoformat(iso_candidate)
        except ValueError:
            pass

    numbers = extract_ints(date_value)
    while numbers:
        try:
            candidate = _datetime_from_parts(numbers[:6])
            return candidate
        except (ValueError, TypeError):
            numbers.pop(0)
    return date_value


def _datetime_from_parts(parts: list[int]) -> datetime:
    if len(parts) < 3:
        raise ValueError("At least year, month, and day are required.")

    safe_parts = parts[:6]
    safe_parts.extend([0] * (6 - len(safe_parts)))
    year, month, day, hour, minute, second = safe_parts
    return datetime(year, month, day, hour, minute, second)


def extract_digits(input_str: str) -> list[str]:
    return DIGIT_REGEX.findall(str(input_str))


def extract_ints(input_str: str) -> list[int]:
    digits = extract_digits(input_str)
    if not digits:
        return []
    return [int(d) for d in digits]


def resolve_property(obj: Any, property_name: str, new_value: Any = _UNSET) -> Any:
    """Resolves a property on an object, optionally setting it to a new value."""
    if not hasattr(obj, property_name):
        if new_value is _UNSET:
            return None
        raise AttributeError(f"Object does not have property '{property_name}'.")

    if new_value is not _UNSET:
        setattr(obj, property_name, new_value)

    if not hasattr(obj, property_name):
        return None

    return cast(Any, getattr(obj, property_name))


def resolve_method(obj: Any, method_name: str, *args: Any, **kwargs: Any) -> Any:
    """Resolves a method on an object and calls it with the given arguments."""
    if not hasattr(obj, method_name):
        raise AttributeError(f"Object does not have method '{method_name}'.")

    method = getattr(obj, method_name)
    if not callable(method):
        raise AttributeError(f"'{method_name}' is not a callable method.")

    return method(*args, **kwargs)
