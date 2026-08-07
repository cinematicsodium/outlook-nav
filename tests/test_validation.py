from datetime import datetime, timezone

import pytest

from outlook import OutlookError
from outlook.validation import validate_datetime, validate_email, validate_paths


def test_email_lists_are_normalized_without_discarding_recipients() -> None:
    assert validate_email(
        "First@Example.com; second@example.com,third@example.com"
    ) == ("first@example.com; second@example.com; third@example.com")
    assert validate_email(["first@example.com", "second@example.com"]) == (
        "first@example.com; second@example.com"
    )


@pytest.mark.parametrize(
    "value", ["invalid-email", "junk valid@example.com", "valid@example.com junk"]
)
def test_invalid_email_rejects_the_complete_input(value: str) -> None:
    with pytest.raises(OutlookError):
        validate_email(value)


def test_datetime_parsing_is_strict_and_does_not_shift_timezones() -> None:
    aware = datetime(2026, 8, 6, 12, tzinfo=timezone.utc)
    assert validate_datetime(aware) is aware
    assert validate_datetime("2026-08-06T12:00:00Z") == aware
    assert validate_datetime("2026-99-99") is None


def test_path_validation_rejects_all_input_when_one_path_is_missing(tmp_path) -> None:
    with pytest.raises(OutlookError):
        validate_paths([tmp_path, tmp_path / "missing"])
