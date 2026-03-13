from __future__ import annotations

from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_address_entry import OlAddressEntry


class OlRecipient(OlItem):
    Address: str
    AddressEntry: OlAddressEntry
    Name: str
    Resolved: bool
    Type: int

    def Delete(self) -> None: ...
    def Resolve(self) -> bool: ...
