from typing import Any

from ..exceptions import OutlookError
from ..models.account import Account


def _load_dispatch() -> Any:
    try:
        from win32com import client  # type: ignore

        return client
    except ImportError as exc:
        raise OutlookError(
            "win32com.client is required. Install pywin32 on Windows."
        ) from exc


def _connect() -> Any:
    try:
        return _load_dispatch().Dispatch("Outlook.Application")
    except OutlookError:
        raise
    except AttributeError as exc:
        raise OutlookError("Unable to connect to Outlook.") from exc


def _open_mapi(app: Any) -> Any:
    try:
        return app.GetNamespace("MAPI")
    except AttributeError as exc:
        raise OutlookError("Unable to open the Outlook MAPI namespace.") from exc


def _select_account(
    accounts: list[Account], value: str | None = None
) -> Account | None:
    if not accounts:
        raise OutlookError("Outlook has no configured email accounts.")
    if value:
        account = next((item for item in accounts if item.matches(value)), None)
        if account is None:
            available = ", ".join(sorted(item.email_address for item in accounts))
            raise OutlookError(
                f"Outlook account not found: {value!r}. Available accounts: {available}."
            )
        return account
    return accounts[0] if len(accounts) == 1 else None
