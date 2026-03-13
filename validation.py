import re
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, TypeVar, overload

from .constants import DIGIT_REGEX, EMAIL_REGEX, TIMEZONE_OFFSET
from .types import LowerStr

T = TypeVar("T")
U = TypeVar("U")


@overload
def ensure_list(
    obj: T | list[T] | set[T] | tuple[T, ...],
    key: None = None,
) -> list[T]: ...


@overload
def ensure_list(
    obj: T | list[T] | set[T] | tuple[T, ...],
    key: Callable[[T], U],
) -> list[U]: ...


def ensure_list(
    obj: T | list[T] | set[T] | tuple[T, ...],
    key: Callable[[T], U] | None = None,
) -> list[T] | list[U]:
    items = list(obj) if isinstance(obj, (list, set, tuple)) else [obj]
    if key is None:
        return items
    return [key(i) for i in items]


def extract_digits(input_str: str) -> list[str]:
    return DIGIT_REGEX.findall(str(input_str))


def parse_digits(digit_list: list[str]) -> list[int]:
    try:
        return [int(d) for d in digit_list]
    except Exception:
        return []


def _extract_datetime_from_str(input_str: str) -> datetime | None:
    digit_list = extract_digits(input_str)
    int_list = parse_digits(digit_list)
    while int_list:
        try:
            return datetime(*int_list)
        except Exception:
            int_list.pop()
    return None


def convert_to_datetime(date_str: str) -> datetime | None:
    """Handles the logic of converting strings to datetime objects."""
    try:
        iso_candidate = date_str.replace("Z", "+00:00")
        return datetime.fromisoformat(iso_candidate)
    except ValueError:
        pass
    return _extract_datetime_from_str(date_str)


def apply_timezone_offset(dt: datetime) -> datetime:
    """Helper to consistently apply the global timezone offset."""
    return dt + TIMEZONE_OFFSET


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
        else convert_to_datetime(str(date_value))
    )
    if dt:
        return apply_timezone_offset(dt)
    return None


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
