from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from protocols import MailItemProtocol
from utils import (
    resolve_method,
    resolve_property,
    validate_datetime,
    validate_email,
    validate_paths,
)

if TYPE_CHECKING:
    from models.folder import Folder


class MailItem:
    def __init__(self, mail_item: MailItemProtocol):
        self.mail_item: MailItemProtocol = mail_item

    @property
    def sender(self) -> str | None:
        """Returns the sender's email address."""
        return cast(str | None, resolve_property(self.mail_item, "SenderEmailAddress"))

    @sender.setter
    def sender(self, value: str) -> None:
        resolve_property(self.mail_item, "SenderEmailAddress", validate_email(value))

    @property
    def sent_on_behalf_of(self) -> str | None:
        """Returns the email address of the person on whose behalf the email was sent."""
        return cast(str | None, resolve_property(self.mail_item, "SentOnBehalfOfName"))

    @sent_on_behalf_of.setter
    def sent_on_behalf_of(self, value: str) -> None:
        resolve_property(self.mail_item, "SentOnBehalfOfName", validate_email(value))

    @property
    def to(self) -> str | None:
        """Returns the recipient's email address."""
        return cast(str | None, resolve_property(self.mail_item, "To"))

    @to.setter
    def to(self, value: str | Iterable[str]) -> None:
        resolve_property(self.mail_item, "To", validate_email(value))

    @property
    def cc(self) -> str | None:
        """Returns the CC recipient's email address."""
        return cast(str | None, resolve_property(self.mail_item, "CC"))

    @cc.setter
    def cc(self, value: str | Iterable[str]) -> None:
        resolve_property(self.mail_item, "CC", validate_email(value))

    @property
    def bcc(self) -> str | None:
        """Returns the BCC recipient's email address."""
        return cast(str | None, resolve_property(self.mail_item, "BCC"))

    @bcc.setter
    def bcc(self, value: str | Iterable[str]) -> None:
        resolve_property(self.mail_item, "BCC", validate_email(value))

    @property
    def recipients(self) -> dict[str, str | None]:
        """Returns a dictionary of all recipients' email addresses."""
        return {
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
        }

    @property
    def subject(self) -> str:
        """Returns the email's subject."""
        return cast(str, resolve_property(self.mail_item, "Subject"))

    @subject.setter
    def subject(self, value: str) -> None:
        resolve_property(self.mail_item, "Subject", value)

    @property
    def body(self) -> str:
        """Returns the email's body."""
        return cast(str, resolve_property(self.mail_item, "Body"))

    @body.setter
    def body(self, value: str) -> None:
        resolve_property(self.mail_item, "Body", value)

    @property
    def html_body(self) -> str:
        """Returns the email's HTML body."""
        return cast(str, resolve_property(self.mail_item, "HTMLBody"))

    @html_body.setter
    def html_body(self, value: str) -> None:
        normalized = value.lower()
        if "<html" not in normalized:
            raise ValueError("HTML body must contain an <html> root element.")
        resolve_property(self.mail_item, "HTMLBody", value)

    @property
    def attachments(self) -> list[str]:
        """Returns a list of attachment file names."""
        attachments = getattr(self.mail_item, "Attachments", None)
        if attachments is not None and getattr(attachments, "Count", 0) > 0:
            return [attachments.Item(i + 1).FileName for i in range(attachments.Count)]
        return []

    @attachments.setter
    def attachments(self, value: str | Path | list[str | Path]) -> None:
        self.add_attachments(value)

    @property
    def scheduled_delivery_time(self) -> datetime | None:
        """Returns the scheduled delivery time of the email, or None if not set."""
        value = resolve_property(self.mail_item, "DeferredDeliveryTime")
        if value is None:
            return None
        normalized = validate_datetime(value)
        if isinstance(normalized, datetime):
            return normalized
        return None

    @scheduled_delivery_time.setter
    def scheduled_delivery_time(self, value: datetime | None) -> None:
        if not isinstance(value, (datetime, type(None))):
            raise ValueError("Scheduled delivery time must be a datetime object or None.")
        resolve_property(self.mail_item, "DeferredDeliveryTime", value)

    @property
    def sent_time(self) -> datetime | Any:
        """Returns the time the email was sent."""
        value = resolve_property(self.mail_item, "SentOn")
        return validate_datetime(value)

    @property
    def received_time(self) -> datetime | Any:
        """Returns the time the email was received."""
        value = resolve_property(self.mail_item, "ReceivedTime")
        return validate_datetime(value)

    @property
    def conversation_id(self) -> str:
        """Returns the email's conversation ID."""
        return cast(str, resolve_property(self.mail_item, "ConversationID"))

    @property
    def size(self) -> int:
        """Returns the size of the email in bytes."""
        value = resolve_property(self.mail_item, "Size")
        try:
            return int(value)
        except (ValueError, TypeError):
            return 0

    @property
    def is_unread(self) -> bool:
        value = resolve_property(self.mail_item, "UnRead")
        return bool(value)

    @is_unread.setter
    def is_unread(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise ValueError("is_unread must be a bool")
        resolve_property(self.mail_item, "UnRead", value)

    def to_dict(self) -> dict[str, Any]:
        return {
            "conversation_id": self.conversation_id,
            "sender": self.sender,
            "sent_on_behalf_of": self.sent_on_behalf_of,
            "recipients": self.recipients,
            "subject": self.subject,
            "body": self.body,
            "attachments": self.attachments,
            "scheduled_delivery_time": self.scheduled_delivery_time,
            "sent_time": self.sent_time,
            "received_time": self.received_time,
            "size": self.size,
        }

    def as_dict(self) -> dict[str, Any]:
        return self.to_dict()

    def send(self) -> None:
        """Sends the email."""
        senders = [self.sender, self.sent_on_behalf_of]

        if not any(senders):
            raise ValueError("Cannot send email without a sender or sent-on-behalf-of address.")
        recipients = [self.to, self.cc, self.bcc]

        if not any(recipients):
            raise ValueError("Cannot send email without at least one recipient.")

        resolve_method(self.mail_item, "Send")

    def save(self) -> None:
        resolve_method(self.mail_item, "Save")

    def delete(self) -> None:
        resolve_method(self.mail_item, "Delete")

    def move(self, destination: Folder) -> MailItem | None:
        from models.folder import Folder

        if not isinstance(destination, Folder):
            raise ValueError("destination must be a Folder")
        moved_item = resolve_method(self.mail_item, "Move", destination.folder)
        if moved_item is None:
            return None
        return MailItem(moved_item)

    def update(self, **fields: Any) -> None:
        allowed_fields = {
            "to",
            "cc",
            "bcc",
            "subject",
            "body",
            "html_body",
            "sender",
            "sent_on_behalf_of",
            "scheduled_delivery_time",
            "is_unread",
        }
        invalid = sorted(set(fields) - allowed_fields)
        if invalid:
            raise ValueError(f"Unsupported update fields: {', '.join(invalid)}")

        for key, value in fields.items():
            setattr(self, key, value)

        self.save()

    def add_attachments(self, path: str | Path | list[str | Path]) -> None:
        attachments = getattr(self.mail_item, "Attachments", None)

        if attachments is None or not hasattr(attachments, "Add"):
            raise AttributeError("Mail item does not support attachments.")

        valid_paths = validate_paths(path)

        for valid_path in valid_paths:
            attachments.Add(str(valid_path))
