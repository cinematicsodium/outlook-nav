from __future__ import annotations

from collections.abc import Iterable
from typing import Any, Protocol


class NamespaceProtocol(Protocol):
    Folders: Iterable[Any]

    def GetDefaultFolder(self, folder_id: int) -> Any: ...

    def CreateRecipient(self, name: str) -> Any: ...


class OutlookApplicationProtocol(Protocol):
    def GetNamespace(self, name: str) -> NamespaceProtocol: ...

    def CreateItem(self, item_type: int) -> Any: ...

    def Quit(self) -> Any: ...


class DispatchModuleProtocol(Protocol):
    def Dispatch(self, app_name: str) -> OutlookApplicationProtocol: ...


class AttachmentItemProtocol(Protocol):
    FileName: str


class AttachmentCollectionProtocol(Protocol):
    Count: int

    def Item(self, index: int) -> AttachmentItemProtocol: ...

    def Add(self, path: str) -> Any: ...


class MailItemProtocol(Protocol):
    Attachments: AttachmentCollectionProtocol | None
    Subject: str
    Body: str
    HTMLBody: str
    UnRead: bool
    SenderEmailAddress: str | None
    SentOnBehalfOfName: str | None
    To: str | None
    CC: str | None
    BCC: str | None

    def Send(self) -> Any: ...

    def Save(self) -> Any: ...

    def Delete(self) -> Any: ...

    def Move(self, folder: Any) -> Any: ...


class FolderProtocol(Protocol):
    Name: str
    Items: Iterable[Any]
    Folders: Iterable[Any]

    def Delete(self) -> Any: ...

    def CopyTo(self, destination: Any) -> Any: ...
