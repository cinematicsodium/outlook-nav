from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING

from .ol_item import OlItem

if TYPE_CHECKING:
    from .ol_account import OlAccount
    from .ol_address_entry import OlAddressEntry
    from .ol_attachment import OlAttachment
    from .ol_collection import OlCollection
    from .ol_folder import OlFolder
    from .ol_inspector import OlInspector
    from .ol_recipient import OlRecipient


class OlMailItem(OlItem):
    """Interface for Outlook Mail Items."""

    AlternateRecipientAllowed: bool
    Attachments: OlCollection[OlAttachment]
    AutoForwarded: bool
    AutoResolvedWinner: bool
    BCC: str
    BillingInformation: str
    Body: str
    BodyFormat: int
    Categories: str
    CC: str
    Companies: str
    ConversationID: str
    ConversationIndex: str
    ConversationTopic: str
    CreationTime: datetime
    DeferredDeliveryTime: datetime
    DeleteAfterSubmit: bool
    DownloadState: int
    EntryID: str
    ExpiryTime: datetime
    FlagRequest: str
    GetInspector: OlInspector
    HTMLBody: str
    Importance: int
    InternetCodepage: int
    IsConflict: bool
    IsMarkedAsTask: bool
    LastModificationTime: datetime
    MarkForDownload: int
    MessageClass: str
    Mileage: str
    NoAging: bool
    OriginatorDeliveryReportRequested: bool
    OutlookInternalVersion: int
    OutlookVersion: str
    Parent: OlFolder
    Permission: int
    PermissionService: int
    PermissionTemplateGuid: str
    ReadReceiptRequested: bool
    ReceivedByEntryID: str
    ReceivedByName: str
    ReceivedOnBehalfOfEntryID: str
    ReceivedOnBehalfOfName: str
    ReceivedTime: datetime
    Recipients: OlCollection[OlRecipient]
    RecipientReassignmentProhibited: bool
    ReminderOverrideDefault: bool
    ReminderPlaySound: bool
    ReminderSet: bool
    ReminderSoundFile: str
    ReminderTime: datetime
    RemoteStatus: int
    ReplyRecipientNames: str
    RetentionExpirationDate: datetime
    RetentionPolicyName: str
    Saved: bool
    Sender: OlAddressEntry
    SenderEmailAddress: str
    SenderEmailType: str
    SenderName: str
    SendUsingAccount: OlAccount
    Sensitivity: int
    Sent: bool
    SentOn: datetime
    SentOnBehalfOfName: str
    Size: int
    Subject: str
    Submitted: bool
    TaskCompletedDate: datetime
    TaskDueDate: datetime
    TaskStartDate: datetime
    TaskSubject: str
    To: str
    ToDoTaskOrdinal: datetime
    UnRead: bool
    VotingOptions: str
    VotingResponse: str

    def Send(self) -> None: ...
    def Save(self) -> None: ...
    def Delete(self) -> None: ...
    def Move(self, folder: OlFolder) -> OlMailItem: ...
    def Display(self, modal: bool = False) -> None: ...
