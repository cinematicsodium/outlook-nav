from __future__ import annotations

from collections.abc import Iterable
from types import SimpleNamespace

import pytest

from _outlook import Account, DefaultFolders, Folder, FolderEnum, ItemType, MailItem


class FakeCollection:
    def __init__(self, items: Iterable[object]) -> None:
        self.items = list(items)
        self.item_calls = 0
        self.sort_calls: list[tuple[str, bool]] = []

    @property
    def Count(self) -> int:
        return len(self.items)

    def Item(self, index: int) -> object:
        self.item_calls += 1
        return self.items[index - 1]

    def Add(self, item: object) -> object:
        self.items.append(item)
        return item

    def Sort(self, field: str, descending: bool) -> None:
        self.sort_calls.append((field, descending))


class RestrictableCollection(FakeCollection):
    restricted: FakeCollection | None = None

    def Restrict(self, query: str) -> FakeCollection:
        assert query == "[UnRead] = True"
        self.restricted = FakeCollection(
            [item for item in self.items if getattr(item, "UnRead", False)]
        )
        return self.restricted


class FakeMailItem:
    Class = ItemType.MAIL_ITEM

    def __init__(self, subject: str, unread: bool = True) -> None:
        self.Subject = subject
        self.UnRead = unread
        self.SenderEmailAddress = "sender@example.com"
        self.SentOnBehalfOfName = "sender@example.com"
        self.To = "receiver@example.com"
        self.CC = self.BCC = ""
        self.Body = "body"
        self.HTMLBody = "<html><body>body</body></html>"
        self.Attachments = FakeCollection([])
        self.DeferredDeliveryTime = self.SentOn = self.ReceivedTime = None
        self.ConversationID = self.ConversationIndex = self.EntryID = subject
        self.Sender = None
        self.SenderName = "Sender"
        self.Parent = None
        self.Size = 1
        self.Recipients = SimpleNamespace(ResolveAll=lambda: True)
        self.saved = 0

    def Save(self) -> None:
        self.saved += 1


class FakeFolder:
    Class = ItemType.FOLDER

    def __init__(
        self,
        name: str,
        items: FakeCollection | None = None,
        subfolders: list[FakeFolder] | None = None,
    ) -> None:
        self.Name = name
        self.FolderPath = name
        self.Items = items or FakeCollection([])
        self.Folders = FakeCollection(subfolders or [])
        for item in self.Items.items:
            if hasattr(item, "Parent"):
                item.Parent = self


class FakeAccount:
    Class = ItemType.ACCOUNT

    def __init__(self, name: str, address: str, root: FakeFolder) -> None:
        self.DisplayName = name
        self.SmtpAddress = address
        self.DeliveryStore = SimpleNamespace(GetRootFolder=lambda: root)


class FakeNamespace:
    Class = ItemType.NAMESPACE

    def __init__(self, folder: FakeFolder) -> None:
        self.folder = folder
        self.calls = 0

    def GetDefaultFolder(self, folder_type: FolderEnum) -> FakeFolder:
        self.calls += 1
        return self.folder


class FlakyNamespace(FakeNamespace):
    def GetDefaultFolder(self, folder_type: FolderEnum) -> FakeFolder:
        if self.calls == 0:
            self.calls += 1
            raise RuntimeError("temporary COM failure")
        return super().GetDefaultFolder(folder_type)


def test_account_finds_nested_folder_and_ignores_empty_path_segments() -> None:
    reports = FakeFolder("Reports")
    inbox = FakeFolder("Inbox", subfolders=[reports])
    account = Account(
        FakeAccount(
            "Primary", "user@example.com", FakeFolder("Root", subfolders=[inbox])
        )
    )

    assert account.matches("USER@EXAMPLE.COM")
    folder = account.find_folder("/Inbox//Reports/")
    assert folder is not None
    assert folder.name == "Reports"
    assert account.find_folder("Inbox/Missing") is None


def test_message_limit_stops_collection_enumeration() -> None:
    collection = FakeCollection([FakeMailItem(str(index)) for index in range(10_000)])
    messages = Folder(FakeFolder("Inbox", items=collection)).list_messages(limit=20)

    assert len(messages) == 20
    assert collection.item_calls == 20
    assert collection.sort_calls == [("[ReceivedTime]", True)]


def test_unread_filter_uses_outlook_restrict_before_enumerating() -> None:
    collection = RestrictableCollection(
        [FakeMailItem("read", unread=False), FakeMailItem("unread")]
    )
    messages = Folder(FakeFolder("Inbox", items=collection)).list_messages(
        limit=1, unread_only=True
    )

    assert [message.subject for message in messages] == ["unread"]
    assert collection.item_calls == 0
    assert collection.restricted is not None
    assert collection.restricted.item_calls == 1


def test_unread_filter_fallback_stops_after_enough_matches() -> None:
    collection = FakeCollection(
        [FakeMailItem(str(index), unread=index % 2 == 0) for index in range(10_000)]
    )
    messages = Folder(FakeFolder("Inbox", items=collection)).list_messages(
        limit=20, unread_only=True
    )

    assert len(messages) == 20
    assert collection.item_calls == 39


def test_get_item_uses_one_direct_com_lookup() -> None:
    collection = FakeCollection([FakeMailItem("first"), FakeMailItem("second")])
    message = Folder(FakeFolder("Inbox", items=collection)).get_item(1)

    assert message is not None
    assert message.subject == "second"
    assert collection.item_calls == 1


def test_default_folder_cache_is_case_insensitive() -> None:
    namespace = FakeNamespace(FakeFolder("Inbox"))
    folders = DefaultFolders(namespace)

    assert folders.get(FolderEnum.INBOX) is folders.get(" Inbox ")
    assert namespace.calls == 1


def test_default_folder_failure_is_not_cached() -> None:
    namespace = FlakyNamespace(FakeFolder("Inbox"))
    folders = DefaultFolders(namespace)

    with pytest.raises(RuntimeError):
        folders.get("inbox")
    assert folders.get("inbox") is not None
    assert namespace.calls == 2


def test_folder_walk_honors_max_depth() -> None:
    leaf = FakeFolder("Leaf")
    child = FakeFolder("Child", subfolders=[leaf])
    root = Folder(FakeFolder("Root", subfolders=[child]))

    assert [entry.path for entry in root.walk(recursive=True, max_depth=1)] == [
        "Root",
        "Root/Child",
    ]


def test_mail_update_rejects_every_key_before_mutating() -> None:
    raw = FakeMailItem("original")
    message = MailItem(raw)

    with pytest.raises(ValueError):
        message.update(subject="changed", legacy_field=True)

    assert raw.Subject == "original"
    assert raw.saved == 0
