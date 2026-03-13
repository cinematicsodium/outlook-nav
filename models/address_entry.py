from __future__ import annotations

from ..enums import ItemType
from ..protocols import OlAddressEntry
from ..utils import get_smtp_address, is_accessible_ol_item


class AddressEntry:
    def __init__(self, ol_address_entry: OlAddressEntry):
        self.ol_address_entry = ol_address_entry

    @classmethod
    def interface_properties(cls):
        return (
            "Address",
            "Name",
            "Type",
            "PropertyAccessor",
            "AddressEntryUserType",
        )

    @classmethod
    def is_accessible_entry(cls, ol_address_entry: OlAddressEntry) -> bool:
        return is_accessible_ol_item(
            ol_address_entry,
            ItemType.ADDRESS_ENTRY,
            cls.interface_properties(),
        )

    @classmethod
    def from_outlook_item(cls, ol_address_entry: OlAddressEntry) -> AddressEntry | None:
        if not cls.is_accessible_entry(ol_address_entry):
            return None
        return cls(ol_address_entry)

    @property
    def name(self) -> str:
        """Returns the display name of the address entry."""
        return str(self.ol_address_entry.Name)

    @property
    def address(self) -> str:
        """Returns the email address of the address entry."""
        return get_smtp_address(self.ol_address_entry) or self.ol_address_entry.Address

    @property
    def property_accessor(self):
        """Returns the PropertyAccessor object of the address entry."""
        return self.ol_address_entry.PropertyAccessor

    @property
    def user_type(self) -> int:
        """Returns the user type of the address entry."""
        return self.ol_address_entry.AddressEntryUserType
