from __future__ import annotations

from collections.abc import Iterator
from datetime import datetime
from typing import Any, Protocol, TypeVar

from outlook.enums import DefaultFolderEnum, OutlookItemClass

T = TypeVar("T", covariant=True)


# ============================================================================
# Base Protocols
# ============================================================================


class OlItem(Protocol):
    Class: int


class OlCollection(OlItem, Protocol[T]):
    """Generic protocol for COM collections."""

    Count: int

    def Item(self, index: int | str) -> T: ...
    def ResolveAll(self) -> T: ...
    def Add(self, obj) -> T: ...
    def __iter__(self) -> Iterator[T]: ...


class OlDispath(Protocol):
    def Dispatch(self, app_name: str) -> OlApplication: ...


# ============================================================================
# Application & Namespace
# ============================================================================


class OlApplication(OlItem):
    def GetNamespace(self, type: str = "MAPI") -> OlNamespace: ...
    def CreateItem(self, item_type: OutlookItemClass | int) -> Any: ...
    def Quit(self) -> None: ...

    # Context manager support for 'with' blocks
    def __enter__(self) -> OlApplication: ...
    def __exit__(self, exc_type, exc_val, exc_tb) -> None: ...


class OlNamespace(OlItem):
    Accounts: OlCollection[OlAccount]
    Folders: OlCollection[OlFolder]

    def GetDefaultFolder(self, folder_type: DefaultFolderEnum | int) -> OlFolder: ...
    def CreateRecipient(self, recipient_name: str) -> OlRecipient: ...
    def GetItemFromID(self, entry_id: str, store_id: Any = None) -> Any: ...


# ============================================================================
# Account & Store
# ============================================================================


class OlAccount(OlItem):
    AccountType: int
    Application: OlApplication
    CurrentUser: OlRecipient
    DeliveryStore: OlStore
    DisplayName: str
    ExchangeMailboxServerName: str
    ExchangeMailboxServerVersion: str
    Parent: OlAccount
    Session: OlNamespace
    SmtpAddress: str
    UserName: str


class OlAccountsCollection(OlCollection): ...


class OlStore(OlItem):
    def GetRootFolder(self) -> OlFolder: ...


# ============================================================================
# Folder
# ============================================================================


class OlFolder(OlItem):
    Name: str
    EntryID: str
    Folders: OlCollection[OlFolder]
    Items: OlCollection[OlMailItem]

    def Delete(self) -> None: ...
    def CopyTo(self, destination: OlFolder) -> OlFolder: ...


# ============================================================================
# Recipients & Addressing
# ============================================================================


class OlExchangeUser(OlItem):
    PrimarySmtpAddress: str
    JobTitle: str
    Department: str


class OlAddressEntry(OlItem):
    Address: str
    Name: str
    Type: str
    PropertyAccessor: OlPropertyAccessor
    AddressEntryUserType: int

    def GetExchangeUser(self) -> OlExchangeUser | None: ...


class OlRecipient(OlItem):
    Address: str
    AddressEntry: OlAddressEntry
    Name: str
    Resolved: bool
    Type: int

    def Delete(self) -> None: ...
    def Resolve(self) -> bool: ...


# ============================================================================
# Mail Items & Components
# ============================================================================


class OlAttachment(OlItem):
    DisplayName: str
    FileName: str
    PathName: str
    Position: int
    Size: int
    Type: int

    def SaveAsFile(self, path: str) -> None: ...
    def Delete(self) -> None: ...


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


# ============================================================================
# Editors & Property Access
# ============================================================================


class OlPropertyAccessor(OlItem):
    def GetProperty(self, schema_name: str) -> Any: ...
    def SetProperty(self, schema_name: str, value: Any) -> None: ...


class OlEditor(OlItem):
    Tables: OlCollection

    def Range(self): ...


class OlInspector(OlItem):
    WordEditor: OlEditor
