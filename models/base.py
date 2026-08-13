from __future__ import annotations

from abc import ABC
from typing import Any, ClassVar, Self

from ..enums import ItemType
from ..utils import is_accessible_ol_item


class ItemModel(ABC):
    """Base wrapper for COM-backed Outlook items.

    Parameters
    ----------
    outlook_item : Any
        Outlook COM object to wrap.

    Raises
    ------
    ValueError
        If the object is inaccessible or is not the expected Outlook type.
    """

    item_type: ClassVar[ItemType]
    required_properties: ClassVar[tuple[str, ...]] = ()
    inaccessible_error_message: ClassVar[str] = (
        "Provided Outlook item is not accessible."
    )

    def __init__(self, outlook_item: Any) -> None:
        if not self.is_accessible(outlook_item):
            raise ValueError(self.inaccessible_error_message)
        self.ol_item = outlook_item

    @classmethod
    def from_outlook_item(cls, outlook_item: Any) -> Self | None:
        """Construct a model from an accessible Outlook item.

        Parameters
        ----------
        outlook_item : Any
            Outlook COM object to wrap.

        Returns
        -------
        ItemModel or None
            A wrapped model, or ``None`` when the item is inaccessible.
        """
        try:
            return cls(outlook_item)
        except ValueError:
            return None

    @classmethod
    def is_accessible(cls, outlook_item: Any) -> bool:
        """Check an Outlook item against the expected class and interface.

        Parameters
        ----------
        outlook_item : Any
            Outlook COM object to inspect.

        Returns
        -------
        bool
            ``True`` when the object can be wrapped by this model.
        """
        return is_accessible_ol_item(
            item=outlook_item,
            target_type=cls.item_type,
            properties=cls.required_properties or None,
        )
