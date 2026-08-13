from __future__ import annotations

import logging
from functools import cached_property
from types import TracebackType

from ..exceptions import OutlookError
from ..protocols import OlApplication, OlMailItem, OlNamespace
from ..services.client import _connect, _open_mapi, _select_account
from ..utils import unpack_collection
from .account import Account
from .mail_item import MailItem

logger = logging.getLogger(__name__)


class Outlook:
    """Connect to Outlook and expose the selected account.

    Parameters
    ----------
    address : str, optional
        Display name or SMTP address of the account to select.
    app : OlApplication, optional
        Existing Outlook application object, primarily for dependency injection.
    mapi : OlNamespace, optional
        Existing MAPI namespace, primarily for dependency injection.

    Raises
    ------
    OutlookError
        If Outlook cannot be opened, has no configured accounts, or ``address``
        does not match an account.
    """

    def __init__(
        self,
        address: str | None = None,
        app: OlApplication | None = None,
        mapi: OlNamespace | None = None,
    ) -> None:
        self._app = app or _connect()
        self._mapi = mapi or _open_mapi(self._app)
        self.account = _select_account(self.accounts, address)
        self.address = self.account.email_address if self.account else None

    def __repr__(self) -> str:
        return f"Outlook(address={self.address!r}, connected={self._app is not None})"

    __str__ = __repr__

    def __enter__(self) -> Outlook:  # ruff: ignore[PYI034]
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        self.close()

    @cached_property
    def accounts(self) -> list[Account]:
        """Return all accounts in the active Outlook profile."""
        return unpack_collection(
            self._require_mapi().Accounts,
            transformer=Account,
        )

    def _require_app(self) -> OlApplication:
        if self._app is None:
            raise OutlookError("Outlook connection is closed or unavailable.")
        return self._app

    def _require_mapi(self) -> OlNamespace:
        if self._mapi is None:
            raise OutlookError("Outlook namespace is closed or unavailable.")
        return self._mapi

    def _require_account(self) -> Account:
        if self.account is None:
            raise OutlookError("No Outlook account is selected.")
        return self.account

    def find_account(self, value: str) -> Account | None:
        """Find an account by display name or SMTP address.

        Parameters
        ----------
        value : str
            Case-insensitive display name or SMTP address.

        Returns
        -------
        Account or None
            The matching account, if present.
        """
        return next(
            (account for account in self.accounts if account.matches(value)),
            None,
        )

    def new_email(self) -> MailItem:
        """Create a new email for the selected account.

        Returns
        -------
        MailItem
            A new unsaved email message.

        Raises
        ------
        OutlookError
            If the connection is closed, no account is selected, or Outlook
            does not return an accessible mail item.
        """
        item: OlMailItem = self._require_app().CreateItem(0)
        item.SentOnBehalfOfName = self.address or self._require_account().email_address
        mail = MailItem.from_outlook_item(item)
        if mail is None:
            raise OutlookError("Unable to create an Outlook email.")
        return mail

    def close(self) -> None:
        """Release held COM references without closing the Outlook process."""
        self.__dict__.pop("accounts", None)
        self.account = None
        self.address = None
        self._mapi = None
        self._app = None
