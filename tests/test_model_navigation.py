from __future__ import annotations

from collections.abc import Iterable
from types import SimpleNamespace

import pytest

from outlook import Account, Folder, MailItem
from outlook.enums import FolderEnum, ItemType


class FakeCollection:
    def __init__(self, items: Iterable[object]) -> None:
        """Initialize a one-indexed fake COM collection.

        Parameters
        ----------
        items : iterable of object
            Initial collection items.

        Returns
        -------
        None
        """
        self.items = list(items)
        self.item_calls = 0
        self.sort_calls: list[tuple[str, bool]] = []

    @property
    def Count(self) -> int:
        """Return the number of items.

        Parameters
        ----------
        None

        Returns
        -------
        int
            Number of items.
        """
        return len(self.items)

    def Item(self, index: int) -> object:
        """Return an item by one-based index.

        Parameters
        ----------
        index : int
            One-based item index.

        Returns
        -------
        object
            Selected item.
        """
        self.item_calls += 1
        return self.items[index - 1]

    def Add(self, item: object) -> object:
        """Append and return an item.

        Parameters
        ----------
        item : object
            Item to append.

        Returns
        -------
        object
            Appended item.
        """
        self.items.append(item)
        return item

    def Sort(self, field: str, descending: bool) -> None:
        """Record a sort request.

        Parameters
        ----------
        field : str
            Requested field name.
        descending : bool
            Requested sort direction.

        Returns
        -------
        None
        """
        self.sort_calls.append((field, descending))


class RestrictableCollection(FakeCollection):
    restricted: FakeCollection | None = None

    def Restrict(self, query: str) -> FakeCollection:
        """Return unread fake messages for the Outlook unread query.

        Parameters
        ----------
        query : str
            Outlook restriction query.

        Returns
        -------
        FakeCollection
            Collection containing unread messages.
        """
        assert query == "[UnRead] = True"
        self.restricted = FakeCollection(
            [item for item in self.items if getattr(item, "UnRead", False)]
        )
        return self.restricted


class FakeMailItem:
    Class = ItemType.MAIL_ITEM

    def __init__(self, subject: str, unread: bool = True) -> None:
        """Initialize a fake mail item.

        Parameters
        ----------
        subject : str
            Message subject.
        unread : bool, default=True
            Initial unread state.

        Returns
        -------
        None
        """
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
        """Record that Outlook saved the item.

        Parameters
        ----------
        None

        Returns
        -------
        None
        """
        self.saved += 1


class FakeFolder:
    Class = ItemType.FOLDER

    def __init__(
        self,
        name: str,
        items: FakeCollection | None = None,
        subfolders: list[FakeFolder] | None = None,
    ) -> None:
        """Initialize a fake folder.

        Parameters
        ----------
        name : str
            Folder name.
        items : FakeCollection, optional
            Folder messages.
        subfolders : list of FakeFolder, optional
            Child folders.

        Returns
        -------
        None
        """
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
        """Initialize a fake account.

        Parameters
        ----------
        name : str
            Account display name.
        address : str
            Account SMTP address.
        root : FakeFolder
            Root folder.

        Returns
        -------
        None
        """
        self.DisplayName = name
        self.SmtpAddress = address
        self.DeliveryStore = FakeStore(root)


class FakeStore:
    def __init__(self, folder: FakeFolder) -> None:
        """Initialize a fake Outlook store.

        Parameters
        ----------
        folder : FakeFolder
            Folder returned by store accessors.

        Returns
        -------
        None
        """
        self.folder = folder
        self.calls = 0

    def GetRootFolder(self) -> FakeFolder:
        """Return the fake store root.

        Parameters
        ----------
        None

        Returns
        -------
        FakeFolder
            Root folder.
        """
        return self.folder

    def GetDefaultFolder(self, folder_type: FolderEnum) -> FakeFolder:
        """Return a fake default folder.

        Parameters
        ----------
        folder_type : FolderEnum
            Requested default folder type.

        Returns
        -------
        FakeFolder
            Configured folder.
        """
        self.calls += 1
        return self.folder


class FlakyStore(FakeStore):
    def GetDefaultFolder(self, folder_type: FolderEnum) -> FakeFolder:
        """Fail once, then return the configured default folder.

        Parameters
        ----------
        folder_type : FolderEnum
            Requested default folder type.

        Returns
        -------
        FakeFolder
            Configured folder after the initial failure.
        """
        if self.calls == 0:
            self.calls += 1
            raise RuntimeError("temporary COM failure")
        return super().GetDefaultFolder(folder_type)


def test_account_finds_nested_folder_and_ignores_empty_path_segments() -> None:
    """Verify nested folder lookup ignores empty path segments.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
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
    """Verify message limits stop collection enumeration early.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    collection = FakeCollection([FakeMailItem(str(index)) for index in range(10_000)])
    messages = Folder(FakeFolder("Inbox", items=collection)).list_messages(limit=20)

    assert len(messages) == 20
    assert collection.item_calls == 20
    assert collection.sort_calls == [("[ReceivedTime]", True)]


def test_unread_filter_uses_outlook_restrict_before_enumerating() -> None:
    """Verify unread filtering delegates to Outlook Restrict.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
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
    """Verify fallback unread filtering stops after enough matches.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    collection = FakeCollection(
        [FakeMailItem(str(index), unread=index % 2 == 0) for index in range(10_000)]
    )
    messages = Folder(FakeFolder("Inbox", items=collection)).list_messages(
        limit=20, unread_only=True
    )

    assert len(messages) == 20
    assert collection.item_calls == 39


def test_get_item_uses_one_direct_com_lookup() -> None:
    """Verify direct message lookup uses one COM collection access.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    collection = FakeCollection([FakeMailItem("first"), FakeMailItem("second")])
    message = Folder(FakeFolder("Inbox", items=collection)).get_item(1)

    assert message is not None
    assert message.subject == "second"
    assert collection.item_calls == 1


def test_account_default_folder_is_cached() -> None:
    """Verify default-folder lookups are cached.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    raw_account = FakeAccount("Primary", "user@example.com", FakeFolder("Inbox"))
    account = Account(raw_account)

    assert account.default_folder(FolderEnum.INBOX) is account.default_folder(
        FolderEnum.INBOX
    )
    assert raw_account.DeliveryStore.calls == 1


def test_account_default_folder_failure_is_not_cached() -> None:
    """Verify failed default-folder lookups are retried.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    raw_account = FakeAccount("Primary", "user@example.com", FakeFolder("Inbox"))
    raw_account.DeliveryStore = FlakyStore(FakeFolder("Inbox"))
    account = Account(raw_account)

    with pytest.raises(RuntimeError):
        account.default_folder(FolderEnum.INBOX)
    assert account.default_folder(FolderEnum.INBOX) is not None
    assert raw_account.DeliveryStore.calls == 2


def test_folder_walk_honors_max_depth() -> None:
    """Verify recursive folder walks stop at the maximum depth.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    leaf = FakeFolder("Leaf")
    child = FakeFolder("Child", subfolders=[leaf])
    root = Folder(FakeFolder("Root", subfolders=[child]))

    assert [entry.path for entry in root.walk(recursive=True, max_depth=1)] == [
        "Root",
        "Root/Child",
    ]


def test_mail_update_rejects_every_key_before_mutating() -> None:
    """Verify invalid mail updates do not mutate the COM item.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    raw = FakeMailItem("original")
    message = MailItem(raw)

    with pytest.raises(ValueError):
        message.update(subject="changed", legacy_field=True)

    assert raw.Subject == "original"
    assert raw.saved == 0
