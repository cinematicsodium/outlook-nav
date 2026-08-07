from __future__ import annotations
from abc import ABC
from typing import Any, ClassVar, Self
from ..enums import ItemType
from ..utils import is_accessible_ol_item


class ItemModel(ABC):
    """Base wrapper for COM-backed Outlook items."""

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
        """Construct a model when the wrapped Outlook item is accessible."""
        try:
            return cls(outlook_item)
        except ValueError:
            return None

    @classmethod
    def is_accessible(cls, outlook_item: Any) -> bool:
        """Validate an Outlook item against the expected class and interface."""
        return is_accessible_ol_item(
            item=outlook_item,
            target_type=cls.item_type,
            properties=cls.required_properties or None,
        )
