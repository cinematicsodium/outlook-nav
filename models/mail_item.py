from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypedDict

from tabulate import tabulate

from ..enums import ItemType
from ..logger import log
from ..protocols import OlMailItem
from ..utils import get_smtp_address, resolve_property, unpack_collection
from ..validation import validate_datetime, validate_email, validate_paths
from .address_entry import AddressEntry
from .node import ItemModel

if TYPE_CHECKING:
    from .folder import Folder


class RecipientInfo(TypedDict):
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
    recipients: list[RecipientInfo]
    subject: str
    body: str
    attachments: list[str]
    scheduled_delivery_time: datetime | None
    sent_time: datetime | None
    received_time: datetime | None
    size: int


class MailItem(ItemModel):
    item_type = ItemType.MAIL_ITEM
    required_properties = (
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
    inaccessible_error_message = "Provided Outlook item is not an accessible mail item."

    def __init__(self, item: OlMailItem):
        super().__init__(item)
        self.ol_mail_item = item
        self._message_table: Any = None

    # ---- Identity / Threading -------------------------------------------------

    @property
    def entry_id(self) -> str:
        """The unique hexadecimal identifier for the Outlook item."""
        return self.ol_mail_item.EntryID or ""

    @property
    def dynamic_uuid(self) -> str:
        """A dynamic unique identifier for the email item, derived from the entry ID."""
        return self.entry_id

    @property
    def conversation_id(self) -> str:
        """A unique string identifying all messages within the same thread."""
        return self.ol_mail_item.ConversationID or ""

    @property
    def conversation_index(self) -> str:
        """A hexadecimal string representing the message's hierarchical position in a thread."""
        return self.ol_mail_item.ConversationIndex or ""

    @property
    def uuid(self) -> str:
        """A unique identifier for the email item, derived from the conversation index."""
        return self.conversation_index

    @property
    def parent_folder(self) -> Folder | None:
        """Returns the Folder containing the mail item."""
        from .folder import Folder

        return Folder.from_outlook_item(self.ol_mail_item.Parent)

    # ---- Sender / Recipients --------------------------------------------------

    @property
    def sender(self) -> AddressEntry | None:
        """Returns an AddressEntry object representing the email sender."""
        return AddressEntry.from_outlook_item(self.ol_mail_item.Sender)

    @property
    def sender_name(self) -> str:
        """Returns the display name of the email sender."""
        return str(self.ol_mail_item.SenderName)

    @property
    def sender_address(self) -> str:
        """Retrieves the sender's email address in lowercase format."""
        if address := get_smtp_address(self.ol_mail_item.Sender):
            return address.lower()
        return str(self.ol_mail_item.SenderEmailAddress or "").lower()

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
    def recipients(self) -> list[RecipientInfo]:
        """Returns a list of RecipientInfo dictionaries for all recipients of the email."""
        recipients = unpack_collection(
            self.ol_mail_item.Recipients, transformer=AddressEntry
        )
        recipient_info = [
            RecipientInfo(name=entry.name, address=entry.email_address)
            for entry in recipients
        ]
        return recipient_info

    # ---- Content ---------------------------------------------------------------

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
        """Gets or sets HTML formatted content of the email body."""
        return self.ol_mail_item.HTMLBody

    @html_body.setter
    def html_body(self, value: str) -> None:
        body = value.lower()
        if body.count("html>") != 2:
            raise ValueError("HTML body must contain an <html> root element.")
        resolve_property(self.ol_mail_item, "HTMLBody", value)

    @property
    def table(self):
        """Gets or sets a Word-based table object within the email body."""
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
        """Gets a list of filenames for all current attachments."""
        ol_attachments = unpack_collection(self.ol_mail_item.Attachments)
        attachment_names = [
            str(att.FileName) for att in ol_attachments if att is not None
        ]
        return attachment_names

    @attachments.setter
    def attachments(self, value: str | Path | list[str | Path]) -> None:
        """Sets the attachments for the email item, replacing any existing attachments."""
        self.add_attachments(value)

    def add_attachments(self, path: str | Path | list[str | Path]) -> None:
        """Attaches one or more files to the email item."""
        valid_paths = validate_paths(path)
        for valid_path in valid_paths:
            self.ol_mail_item.Attachments.Add(str(valid_path))

    # ---- Timing / Status -------------------------------------------------------

    @property
    def scheduled_delivery_time(self) -> datetime | None:
        """Gets or sets the date and time for deferred delivery."""
        return self.ol_mail_item.DeferredDeliveryTime

    @scheduled_delivery_time.setter
    def scheduled_delivery_time(self, value: datetime | None) -> None:
        self.ol_mail_item.DeferredDeliveryTime = validate_datetime(value)

    @property
    def sent_time(self) -> datetime | None:
        """Returns when the email message was sent."""
        return self.ol_mail_item.SentOn

    @property
    def received_time(self) -> datetime | None:
        """Returns when the email message was received."""
        return self.ol_mail_item.ReceivedTime

    @property
    def size(self) -> int:
        """Retrieves the total size of the email item in bytes."""
        return self.ol_mail_item.Size

    @property
    def is_unread(self) -> bool:
        """Gets or sets whether the email message is unread."""
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
    def _dispatch(self) -> str:
        """Internal property for primary timestamp or draft status."""
        return str(self.sent_time or self.received_time or "Draft")

    # ---- Actions ----------------------------------------------------------------

    def display(self) -> None:
        self.ol_mail_item.Display()

    def send(self) -> None:
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
            self.display()
            self.ol_mail_item.Send()
        except Exception as e:
            log.error(
                f"Outlook COM Error: Failed to transmit email '{self.subject}'. Details: {e}"
            )

    def save(self) -> None:
        try:
            self.ol_mail_item.Save()
        except Exception as e:
            log.error(f"Failed to save changes to email '{self.subject}': {e}")

    def delete(self) -> None:
        try:
            self.ol_mail_item.Delete()
        except Exception as e:
            log.error(f"Failed to delete email '{self.subject}': {e}")

    def move_to(self, destination: Folder) -> MailItem | None:

        if not isinstance(destination, Folder):
            raise TypeError(
                f"Invalid destination type for 'move_to'. Expected an instance of "
                f"'{Folder.__module__}.Folder', but received '{type(destination).__name__}'."
            )

        try:
            moved_item = self.ol_mail_item.Move(destination._ol_folder_item)
            if moved_item and moved_item.Class == ItemType.MAIL_ITEM:
                return MailItem.from_outlook_item(moved_item)
        except Exception as e:
            log.error(
                f"Failed to move email '{self.subject}' to folder '{destination.name}': {e}"
            )
        return None

    def update(self, **kwargs: Any) -> None:
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
        invalid = sorted(set(kwargs) - allowed_fields)
        if invalid:
            raise ValueError(f"Unsupported update fields: {', '.join(invalid)}")

        for key, value in kwargs.items():
            setattr(self, key, value)

        self.save()

    def get(self, key: Any, default=None):
        if hasattr(self, key):
            return getattr(self, key)
        if hasattr(self.ol_mail_item, key):
            return getattr(self.ol_mail_item, key)
        return default

    # ---- Serialization ----------------------------------------------------------

    def _to_dict(self) -> MailItemData:
        return MailItemData(
            conversation_index=self.conversation_index,
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
            sent_time=self.sent_time,
            received_time=self.received_time,
            size=self.size,
        )

    def as_dict(self) -> MailItemData:
        return self._to_dict()

    def _to_table(self) -> str:
        fmtd_body = self._fmt_body(self.body or self.html_body or "")

        attachments = f"{len(self.attachments)} files(s)"
        parent_folder = self.parent_folder.name if self.parent_folder else ""

        table_data = (
            ("Subject", self.subject),
            ("Sender", self.sender_address),
            ("To", self.to),
            ("CC", self.cc),
            ("BCC", self.bcc),
            ("Sent On Behalf Of", self.sent_on_behalf_of_address),
            ("Sent Time", self._fmt_dt(self.sent_time)),
            ("Received Time", self._fmt_dt(self.received_time)),
            ("Body", fmtd_body),
            ("Attachments", attachments),
            ("Size", self.size),
            ("Parent Folder", parent_folder),
            ("Conversation ID", self.conversation_id),
            ("Conversation Index", self.conversation_index),
            ("Entry ID", self.entry_id),
        )
        table = tabulate(table_data, tablefmt="grid", maxcolwidths=80)
        return table

    def as_table(self) -> str:
        return self._to_table()

    # ---- Internals / Dunder ----------------------------------------------------

    def _fmt_body(self, body: str):
        if not body or not isinstance(body, str):
            return ""

        body = body.replace("\r", "\n")
        lines = body.splitlines()
        fmtd_lines = [line.strip() for line in lines if line.strip()]
        return "\n".join(fmtd_lines)

    def _fmt_dt(self, dt: datetime | None) -> str:
        if dt is None:
            return ""
        try:
            if abs(dt.year - datetime.now().year) >= 1000:
                return ""
            return str(dt)
        except Exception:
            return ""

    def __str__(self) -> str:
        return f"[{self._dispatch}] {self.sender_address}: {self.subject}"

    def __repr__(self) -> str:
        return (
            f"MailItem(subject={self.subject!r}, sender_address={self.sender_address!r}, "
            f"to={self.to!r}, sent_time={self.sent_time!r})"
        )
