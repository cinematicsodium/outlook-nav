from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem
from .ol_namespace import OlNamespace
from .ol_recipient import OlRecipient
from .ol_store import OlStore

if TYPE_CHECKING:
    from .ol_application import OlApplication


class OlAccount(OlItem):
    AccountType: int
    Application: OlApplication
    CurrentUser: OlRecipient
    DeliveryStore: OlStore
    DisplayName: str
    ExchangeMailboxServerName: str
    ExchangeMailboxServerVersion: str
    Parent: OlAccount
    Session: OlNamespace
    SmtpAddress: str
    UserName: str
