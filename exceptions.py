class OutlookError(Exception):
    """Base exception for outlook-nav."""


class OutlookConnectionError(OutlookError, RuntimeError):
    """Raised when an Outlook COM connection cannot be established or used."""


class OutlookValidationError(OutlookError, ValueError):
    """Base exception for validation errors."""


class EmailValidationError(OutlookValidationError):
    """Raised when one or more email addresses are invalid."""


class PathValidationError(OutlookValidationError):
    """Raised when one or more filesystem paths are invalid."""
