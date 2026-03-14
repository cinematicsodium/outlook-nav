from __future__ import annotations

from ..enums import ItemType
from ..protocols import OlAddressEntry
from ..utils import get_smtp_address
from .node import ItemModel


class AddressEntry(ItemModel):
    item_type = ItemType.ADDRESS_ENTRY
    required_properties = (
        "Address",
        "Name",
        "Type",
        "PropertyAccessor",
        "AddressEntryUserType",
    )
    inaccessible_error_message = (
        "Provided Outlook item is not an accessible address entry."
    )

    def __init__(self, ol_address_entry: OlAddressEntry):
        super().__init__(ol_address_entry)
        self.ol_address_entry = ol_address_entry

    @classmethod
    def interface_properties(cls) -> tuple[str, ...]:
        return cls.required_properties

    @classmethod
    def is_accessible_entry(cls, ol_address_entry: OlAddressEntry) -> bool:
        return cls.is_accessible(ol_address_entry)

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
