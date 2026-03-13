from __future__ import annotations

from .ol_item import OlItem


class OlExchangeUser(OlItem):
    PrimarySmtpAddress: str
    JobTitle: str
    Department: str
