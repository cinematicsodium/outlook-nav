from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime

from pathlib import Path
from typing import TYPE_CHECKING, Any, TypedDict

from tabulate import tabulate

from ..enums import ItemType
from ..logger import log
from ..protocols import OlMailItem
from ..utils import (
    get_smtp_address,
    is_builtin_class,
    resolve_property,
    unpack_collection,
)
from ..validation import validate_datetime, validate_email, validate_paths
from .address_entry import AddressEntry
from .node import ItemModel
from ..type_defs import TableFmtStr, StrPath

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
    byte_size: int


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
        self._table_item: Any = None
        self._table_text: str = ""
        self._as_table_output: str = ""

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
        body_lower = value.lower()
        if body_lower.count("html>") != 2:
            raise ValueError("HTML body must contain an <html> root element.")
        resolve_property(self.ol_mail_item, "HTMLBody", value)

    @property
    def table(self):
        """Gets or sets a Word-based table object within the email body."""
        return self._table_item

    @table.setter
    def table(self, value):
        word_editor = self.ol_mail_item.GetInspector.WordEditor
        editor_range = word_editor.Range()
        table = word_editor.Tables.Add(editor_range, 1, 1)
        table.Borders.Enable = False
        table.PreferredWidthType = 1
        table.PreferredWidth = 8 * 72
        self._table_text = value
        table_item = table.Cell(1, 1).Range.Text = value
        self._table_item = table_item

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
    def size_bytes(self) -> int:
        """Retrieves the total size of the email item in bytes."""
        return self.ol_mail_item.Size

    @property
    def size_megabytes(self) -> float:
        return round(self.size_bytes / (1024**2), 2)

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

        resolved_all = self.ol_mail_item.Recipients.ResolveAll()
        if resolved_all is False:
            log.warning("Sending email '%s' with unresolved recipients", self.subject)

        if self.table and not (self.body or self.html_body):
            self.body = self.table

        if self._as_table_output == "":
            self._as_table_output = self.as_table()

        try:
            self.display()
            self.ol_mail_item.Send()
        except Exception:
            log.exception("Failed to send email '%s'", self.subject)

    def save(self) -> None:
        try:
            self.ol_mail_item.Save()
        except Exception:
            log.exception("Failed to save email '%s'", self.subject)

    def export(self, output_path: StrPath) -> None:
        try:
            target = Path(output_path).expanduser().resolve()
            target.parent.mkdir(parents=True, exist_ok=True)
            self.ol_mail_item.SaveAs(str(target))
        except Exception:
            log.error("Failed to export email to '%s'", output_path)

    def delete(self) -> None:
        try:
            self.ol_mail_item.Delete()
        except Exception:
            log.exception("Failed to delete email '%s'", self.subject)

    def move_to(self, destination: Folder) -> MailItem | None:

        if not isinstance(destination, Folder):
            raise TypeError(
                f"Invalid destination type for 'move_to'. Expected an instance of "
                f"'{Folder.__module__}.Folder', but received '{type(destination).__name__}'."
            )

        try:
            moved_item = self.ol_mail_item.Move(destination._ol_folder_item)
            if moved_item is None:
                log.warning(
                    "Move of email '%s' to folder '%s' returned no item",
                    self.subject,
                    destination.name,
                )
                return None
            if moved_item.Class == ItemType.MAIL_ITEM:
                return MailItem.from_outlook_item(moved_item)
            log.warning(
                "Move of email '%s' to folder '%s' returned unexpected item class %r",
                self.subject,
                destination.name,
                getattr(moved_item, "Class", None),
            )
        except Exception:
            log.exception(
                "Failed to move email '%s' to folder '%s'",
                self.subject,
                destination.name,
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
        data = MailItemData(
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
            byte_size=self.size_bytes,
        )
        cleaned = {key: self._format_val(val) for key, val in data.items()}
        return cleaned

    def as_dict(self) -> MailItemData:
        return self._to_dict()

    def as_table(
        self,
        table_format: TableFmtStr = "grid",
        table_width: int | None = 100,
        body_char_limit: int | None = 80,
    ) -> str:
        if (output := self._as_table_output) != "":
            self._as_table_output = ""
            return output

        parent_folder_name = self.parent_folder.name if self.parent_folder else ""

        attachments = f"{len(self.attachments)} files(s)"

        body_selection = self.body or self._table_text or ""
        body = self._format_body(body_selection, body_char_limit)

        size = f"{self.size_bytes} bytes ({self.size_megabytes} MB)"

        scheduled = self._format_datetime(self.scheduled_delivery_time)
        sent = self._format_datetime(self.sent_time)
        received = self._format_datetime(self.received_time)

        data = {
            "Subject": self.subject,
            "Sender": self.sender_address,
            "To": self.to,
            "CC": self.cc,
            "BCC": self.bcc,
            "Sent On Behalf Of": self.sent_on_behalf_of_address,
            "Scheduled Delivery Time": scheduled,
            "Sent Time": sent,
            "Received Time": received,
            "Body": body,
            "Attachments": attachments,
            "Size": size,
            "Parent Folder": parent_folder_name,
            "Conversation ID": self.conversation_id,
            "Conversation Index": self.conversation_index,
            "Entry ID": self.entry_id,
        }
        items = data.items()
        table = tabulate(items, tablefmt=table_format, maxcolwidths=table_width).strip()
        self._as_table_output = table
        return table

    # ---- Internals / Dunder ----------------------------------------------------

    def _format_val(self, val: Any):
        if isinstance(val, str):
            lines = [
                " ".join(line.strip().split())
                for line in val.splitlines()
                if line.strip()
            ]
            return "\n".join(lines)
        return val if is_builtin_class(val) else str(val)

    def _format_body(self, body: str, char_limit: int | None = None):
        if not body or not isinstance(body, str):
            return ""
        formatted = self._format_val(body)
        formatted = formatted if isinstance(formatted, str) else body
        return formatted[:char_limit] if isinstance(char_limit, int) else formatted

    def _format_datetime(self, datetime_value: datetime | None) -> str:
        if datetime_value is None:
            return ""
        try:
            if abs(datetime_value.year - datetime.now().year) >= 100:
                return ""
            return str(datetime_value)
        except Exception:
            return ""

    def __str__(self) -> str:
        return f"[{self._dispatch}] {self.sender_address}: {self.subject}"

    def __repr__(self) -> str:
        return (
            f"MailItem(subject={self.subject!r}, sender_address={self.sender_address!r}, "
            f"to={self.to!r}, sent_time={self.sent_time!r})"
        )
