from __future__ import annotations

from .ol_exchange_user import OlExchangeUser
from .ol_item import OlItem
from .ol_property_accessor import OlPropertyAccessor


class OlAddressEntry(OlItem):
    Address: str
    Name: str
    Type: str
    PropertyAccessor: OlPropertyAccessor
    AddressEntryUserType: int

    def GetExchangeUser(self) -> OlExchangeUser | None: ...
