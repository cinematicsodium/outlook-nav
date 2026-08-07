import re
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

from .exceptions import OutlookError

_EMAIL = re.compile(
    r"[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,63}",
    re.IGNORECASE,
)
_RECIPIENT_SEPARATOR = re.compile(r"\s*[;,]\s*")


def validate_datetime(value: Any) -> datetime | None:
    """Return a datetime or None; strings must use ISO 8601 syntax."""
    if value is None or isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def validate_email(emails: str | Iterable[str] | None) -> str:
    """Validate recipients and return Outlook's semicolon-delimited form."""
    if not emails:
        return ""
    values = [emails] if isinstance(emails, str) else list(emails)
    parsed = [
        address.strip().lower()
        for value in values
        for address in _RECIPIENT_SEPARATOR.split(str(value))
        if address.strip()
    ]
    invalid = [address for address in parsed if _EMAIL.fullmatch(address) is None]
    if invalid:
        raise OutlookError(f"Invalid emails: {invalid}")
    return "; ".join(parsed)


def validate_paths(paths: str | Path | Iterable[str | Path]) -> list[Path]:
    """Resolve paths and reject the complete input when any path is invalid."""
    values = [paths] if isinstance(paths, (str, Path)) else list(paths)
    valid: list[Path] = []
    errors: list[str] = []
    for value in values:
        if not isinstance(value, (str, Path)):
            errors.append(f"Invalid path: {value} (type: {type(value)})")
            continue
        path = Path(value).expanduser().resolve()
        if path.exists():
            valid.append(path)
        else:
            errors.append(f"Path does not exist: {path}")
    if errors:
        raise OutlookError("Path validation errors:\n" + "\n".join(errors))
    return valid
