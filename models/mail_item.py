from __future__ import annotations

import logging
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypedDict

from outlook.constants import OutlookItemClass
from outlook.protocols import OlMailItem
from outlook.utils import (
    get_smtp_address,
    is_valid_ol_item,
    resolve_property,
    validate_datetime,
    validate_email,
    validate_paths,
)

log = logging.getLogger(__name__)

if TYPE_CHECKING:
    from outlook.models.folder import Folder


class RecipientsData(TypedDict):
    name: str
    address: str


class MailItemData(TypedDict):
    conversation_index: str
    conversation_id: str
    sender_address: str
    sent_on_behalf_of_address: str
    to: str
    cc: str
    bcc: str
    recipients: RecipientsData
    subject: str
    body: str
    attachments: list[str]
    scheduled_delivery_time: datetime
    sent_time: datetime
    received_time: datetime
    size: int


class MailItem:
    def __init__(self, item: Any):
        self.ol_mail_item: OlMailItem = item
        self._message_table: Any = None

    @classmethod
    def from_outlook_item(cls, item: Any) -> MailItem | None:
        if not is_valid_ol_item(
            item=item,
            target_type=OutlookItemClass.MAIL_ITEM,
            properties=MailItem.interface_properties(),
        ):
            return None

        return cls(item)

    @staticmethod
    def interface_properties():
        return (
            "SenderEmailAddress",
            "SentOnBehalfOfName",
            "To",
            "CC",
            "BCC",
            "Subject",
            "Body",
            "HTMLBody",
            "Attachments",
            "DeferredDeliveryTime",
            "SentOn",
            "ReceivedTime",
            "ConversationID",
        )

    @property
    def entry_id(self):
        """
        The unique hexadecimal identifier for the Outlook item.

        Note:
            This ID is storage-dependent. Moving the item between different
            stores (e.g., from a mailbox to a .pst file) will change the ID.
        """
        return self.ol_mail_item.EntryID

    @property
    def conversation_id(self):
        """
        A unique string identifying all messages within the same thread.

        Used to relate replies and forwards across different folders.
        """
        return self.ol_mail_item.ConversationID

    @property
    def conversation_index(self):
        """
        A hexadecimal string representing the message's hierarchical position in a thread.

        The root message is `44` characters long (`22` bytes), with `10` characters
        (`5` bytes) appended for each subsequent reply level.
        """
        return self.ol_mail_item.ConversationIndex

    @property
    def parent_folder(self):
        """returns the Folder containing the mail item."""
        from .folder import Folder

        return Folder.from_outlook_item(self.ol_mail_item.Parent)

    @property
    def sender_name(self) -> str:
        """returns the display name of the email sender."""
        return str(self.ol_mail_item.SenderName)

    @property
    def sender_address(self) -> str:
        """Retrieves the sender's email address in lowercase format."""
        addresses = (
            get_smtp_address(self.ol_mail_item.Sender),
            self.ol_mail_item.SenderEmailAddress,
        )
        if address := get_smtp_address(self.ol_mail_item.Sender):
            return address.lower()
        if address := self.ol_mail_item.SenderEmailAddress:
            return address.lower()
        return ""

    @property
    def sent_on_behalf_of_address(self) -> str:
        """Gets or sets the name or address of the person the email is sent on behalf of."""
        return self.ol_mail_item.SentOnBehalfOfName

    @sent_on_behalf_of_address.setter
    def sent_on_behalf_of_address(self, value: str) -> None:
        self.ol_mail_item.SentOnBehalfOfName = validate_email(value)

    @property
    def to(self) -> str:
        """Gets or sets the primary recipients of the email message."""
        return self.ol_mail_item.To

    @to.setter
    def to(self, value: str | Iterable[str]) -> None:
        self.ol_mail_item.To = validate_email(value)

    @property
    def cc(self) -> str:
        """Gets or sets the carbon copy recipients of the email message."""
        return self.ol_mail_item.CC

    @cc.setter
    def cc(self, value: str | Iterable[str]) -> None:
        self.ol_mail_item.CC = validate_email(value)

    @property
    def bcc(self) -> str:
        """Gets or sets the blind carbon copy recipients of the email message."""
        return self.ol_mail_item.BCC

    @bcc.setter
    def bcc(self, value: str | Iterable[str]) -> None:
        self.ol_mail_item.BCC = validate_email(value)

    @property
    def recipients(self):
        """Returns a dictionary mapping each recipient's email address (str) to a nested dictionary containing their name and address details."""
        recipients_data: dict[str, RecipientsData] = {}
        try:
            recipients = self.ol_mail_item.Recipients
            if not recipients.Count:
                return recipients_data

            for i in range(recipients.Count + 1):

                recipient = recipients.Item(i)
                recipient_name = recipient.Name
                recipient_address = ""

                if smtp_address := get_smtp_address(recipient.AddressEntry):
                    recipient_address = smtp_address

                recipients_data[recipient_address] = {
                    "name": recipient_name,
                    "address": recipient_address,
                }

        except Exception:
            pass
        return recipients_data

    @property
    def subject(self) -> str:
        """Gets or sets the subject line for the email message."""
        return self.ol_mail_item.Subject

    @subject.setter
    def subject(self, value: str) -> None:
        self.ol_mail_item.Subject = value

    @property
    def body(self) -> str:
        """Gets or sets the plain text content of the email body."""
        return self.ol_mail_item.Body

    @body.setter
    def body(self, value: str) -> None:
        self.ol_mail_item.Body = value

    @property
    def html_body(self) -> str:
        """Gets or sets the HTML formatted content of the email body with validation for root tags."""
        return self._fmt_body(self.ol_mail_item.HTMLBody)

    @html_body.setter
    def html_body(self, value: str) -> None:
        body = value.lower()
        if body.count("html>") != 2:
            raise ValueError("HTML body must contain an <html> root element.")
        resolve_property(self.ol_mail_item, "HTMLBody", value)

    @property
    def table(self):
        """Gets or sets a Word-based table object within the email body using the internal inspector."""
        return self._message_table

    @table.setter
    def table(self, value):
        word_editor = self.ol_mail_item.GetInspector.WordEditor
        editor_range = word_editor.Range()
        table = word_editor.Tables.Add(editor_range, 1, 1)
        table.Borders.Enable = False
        table.PreferredWidthType = 1
        table.PreferredWidth = 8 * 72
        self._message_table = table.Cell(1, 1).Range.Text = value

    @property
    def attachments(self) -> list[str]:
        """Gets a list of filenames for all current attachments or adds new attachments to the message."""
        attachments = self.ol_mail_item.Attachments
        results = []
        if attachments is not None and attachments.Count > 0:
            for i in range(attachments.Count):
                try:
                    results.append(attachments.Item(i + 1).FileName)
                except Exception:
                    pass
        return results

    @attachments.setter
    def attachments(self, value: str | Path | list[str | Path]) -> None:
        self.add_attachments(value)

    def add_attachments(self, path: str | Path | list[str | Path]) -> None:
        """Attaches one or more files to the email item after validating the provided paths."""
        valid_paths = validate_paths(path)
        for valid_path in valid_paths:
            self.ol_mail_item.Attachments.Add(str(valid_path))

    @property
    def scheduled_delivery_time(self) -> datetime | None:
        """Gets or sets the date and time for deferred delivery of the email message."""
        return self.ol_mail_item.DeferredDeliveryTime

    @scheduled_delivery_time.setter
    def scheduled_delivery_time(self, value: datetime | None) -> None:
        validated = validate_datetime(value)
        self.ol_mail_item.DeferredDeliveryTime = validated

    @property
    def sent_time(self) -> datetime | None:
        """Returns the timestamp indicating when the email message was sent."""
        return self.ol_mail_item.SentOn

    @property
    def received_time(self) -> datetime | None:
        """Returns the timestamp indicating when the email message was received."""
        return self.ol_mail_item.ReceivedTime

    @property
    def size(self) -> int:
        """Retrieves the total size of the email item in bytes."""
        return self.ol_mail_item.Size

    @property
    def is_unread(self) -> bool:
        """Gets or sets a boolean indicating whether the email message is marked as unread."""
        return self.ol_mail_item.UnRead

    @is_unread.setter
    def is_unread(self, value: bool) -> None:
        if not isinstance(value, bool):
            raise ValueError(
                f"The 'is_unread' property must be a boolean (True or False), "
                f"but received {type(value).__name__}: {value!r}"
            )
        self.ol_mail_item.UnRead = value

    @property
    def _dispatch(self):
        """Internal property that returns the primary timestamp or draft status of the email."""
        return self.sent_time or self.received_time or "Draft"

    def display(self):
        """Opens the Outlook inspector window to display the email item to the user."""
        self.ol_mail_item.Display()

    def send(self) -> None:
        """Validates sender and recipient requirements before transmitting the email via Outlook."""
        senders = [self.sender_address, self.sent_on_behalf_of_address]

        if not any(senders):
            raise ValueError(
                f"Unable to send email '{self.subject}': No sender identified. "
                f"Both 'sender_address' and 'sent_on_behalf_of_address' are empty."
            )

        recipients = [self.to, self.cc, self.bcc]

        if not any(recipients):
            raise ValueError(
                f"Unable to send email '{self.subject}': At least one recipient must be "
                f"provided in the 'to', 'cc', or 'bcc' fields."
            )

        self.ol_mail_item.Recipients.ResolveAll()

        if self.table and not (self.body or self.html_body):
            self.body = self.table

        try:
            self.ol_mail_item.Send()
        except Exception as e:
            log.error(
                f"Outlook COM Error: Failed to transmit email '{self.subject}'. Details: {e}"
            )

        return

    def save(self) -> None:
        """Saves any modifications made to the email item in its current folder."""
        try:
            self.ol_mail_item.Save()
        except Exception as e:
            log.error(f"Failed to save changes to email '{self.subject}': {e}")

    def delete(self) -> None:
        """Permanently removes the email item from its current folder."""
        try:
            self.ol_mail_item.Delete()
        except Exception as e:
            log.error(f"Failed to delete email '{self.subject}': {e}")

    def move_to(self, destination: Folder) -> MailItem | None:
        """Transfers the email item to a new folder and returns the moved item instance."""
        from outlook.models.folder import Folder

        if not isinstance(destination, Folder):
            raise TypeError(
                f"Invalid destination type for 'move_to'. Expected an instance of "
                f"'{Folder.__module__}.Folder', but received '{type(destination).__name__}'."
            )
        try:
            moved_item = self.ol_mail_item.Move(destination._ol_folder_item)
            if moved_item and moved_item.Class == OutlookItemClass.MAIL_ITEM:
                return MailItem.from_outlook_item(moved_item)
        except Exception as e:
            log.error(
                f"Failed to move email '{self.subject}' to folder '{destination.name}': {e}"
            )

    def update(self, **fields: Any) -> None:
        """Performs a bulk update on multiple email fields and saves the changes."""
        allowed_fields = {
            "to",
            "cc",
            "bcc",
            "subject",
            "body",
            "html_body",
            "sender",
            "sent_on_behalf_of_address",
            "scheduled_delivery_time",
            "is_unread",
        }
        invalid = sorted(set(fields) - allowed_fields)
        if invalid:
            raise ValueError(f"Unsupported update fields: {', '.join(invalid)}")

        for key, value in fields.items():
            setattr(self, key, value)

        self.save()

    def get(self, key: Any, default=None):
        """Safely retrieves a property value from either the wrapper class or the underlying Outlook object."""
        if hasattr(self, key):
            return getattr(self, key)
        if hasattr(self.ol_mail_item, key):
            return getattr(self.ol_mail_item, key)
        return default

    def to_dict(self) -> MailItemData:
        """Serializes the core properties of the email item into a structured dictionary format."""
        return MailItemData(
            conversation_index=self.entry_id,
            conversation_id=self.conversation_id,
            sender_address=self.sender_address,
            sent_on_behalf_of_address=self.sent_on_behalf_of_address,
            to=self.to,
            cc=self.cc,
            bcc=self.bcc,
            recipients=self.recipients,
            subject=self.subject,
            body=self.body,
            attachments=self.attachments,
            scheduled_delivery_time=self.scheduled_delivery_time,
            received_time=self.received_time,
            size=self.size,
        )

    def as_dict(self) -> MailItemData:
        """Converts the email item properties into a dictionary format by calling the `.to_dict()` method."""
        return self.to_dict()

    def _fmt_body(self, body: str):
        if not body or not isinstance(body, str):
            return ""

        body = body.replace("\r", "\n")
        lines = body.splitlines()
        fmtd_lines = [line.strip() for line in lines if line.strip()]
        fmtd_body = "\n".join(fmtd_lines)

        return fmtd_body

    def __str__(self) -> str:
        return f"[{self._dispatch}] {self.sender_address}: {self.subject}"

    def __repr__(self) -> str:

        return (
            f"parent_folder={self.parent_folder.name}, "
            f"dispatch={self._dispatch}, "
            f"subject={self.subject}, "
            f"sender={self.sender_address})"
        )
