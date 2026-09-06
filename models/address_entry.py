from __future__ import annotations

from ..enums import ItemType
from ..protocols import OlAddressEntry
from ..utils import get_smtp_address
from .base import ItemModel


class AddressEntry(ItemModel):
    """Represent an Outlook address entry.

    Parameters
    ----------
    ol_address_entry : OlAddressEntry
        Outlook address entry COM object to wrap.
    """

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
        """Initialize an address-entry wrapper.

        Parameters
        ----------
        ol_address_entry : OlAddressEntry
            Outlook address-entry COM object.

        Returns
        -------
        None
        """
        super().__init__(ol_address_entry)
        self._ol_address_entry = ol_address_entry

    @property
    def name(self) -> str:
        """Returns the display name of the address entry."""
        return str(self._ol_address_entry.Name)

    @property
    def email_address(self) -> str:
        """Returns the email address of the address entry."""
        if address := get_smtp_address(self._ol_address_entry):
            return address
        return self._ol_address_entry.Address or ""

    @property
    def property_accessor(self):
        """Returns the PropertyAccessor object of the address entry."""
        return self._ol_address_entry.PropertyAccessor

    @property
    def user_type(self) -> int:
        """Returns the user type of the address entry."""
        return self._ol_address_entry.AddressEntryUserType
