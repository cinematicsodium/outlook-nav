import importlib
import sys
import types
import unittest
from pathlib import Path


def _bootstrap_outlook_package() -> None:
    if "outlook" in sys.modules:
        return

    root = Path(__file__).resolve().parents[1]
    pkg = types.ModuleType("outlook")
    pkg.__path__ = [str(root)]
    sys.modules["outlook"] = pkg


_bootstrap_outlook_package()

enums = importlib.import_module("outlook.enums")
account_module = importlib.import_module("outlook.models.account")
default_folders_module = importlib.import_module("outlook.models.default_folders")
folder_module = importlib.import_module("outlook.models.folder")

Account = account_module.Account
DefaultFolders = default_folders_module.DefaultFolders
Folder = folder_module.Folder
FolderType = enums.FolderType
ItemType = enums.ItemType


class FakeCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = list(items)

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> object:
        return self._items[index - 1]

    def Add(self, item: object) -> object:
        self._items.append(item)
        return item


class FakeMailItem:
    Class = ItemType.MAIL_ITEM

    def __init__(self, subject: str, unread: bool = True) -> None:
        self.SenderEmailAddress = "sender@example.com"
        self.SentOnBehalfOfName = ""
        self.To = "receiver@example.com"
        self.CC = ""
        self.BCC = ""
        self.Subject = subject
        self.Body = f"Body for {subject}"
        self.HTMLBody = "<html><body>Body</body></html>"
        self.Attachments = FakeCollection([])
        self.DeferredDeliveryTime = None
        self.SentOn = None
        self.ReceivedTime = None
        self.ConversationID = f"conversation-{subject}"
        self.ConversationIndex = f"index-{subject}"
        self.EntryID = f"entry-{subject}"
        self.Sender = None
        self.SenderName = "Sender"
        self.Parent = None
        self.Size = 1
        self.UnRead = unread
        self.Recipients = types.SimpleNamespace(ResolveAll=lambda: True)


class FakeFolder:
    Class = ItemType.FOLDER

    def __init__(
        self,
        name: str,
        items: list[object] | None = None,
        subfolders: list["FakeFolder"] | None = None,
    ) -> None:
        self.Name = name
        self.Items = FakeCollection(items or [])
        self.Folders = FakeCollection(subfolders or [])
        self.deleted = False

        for item in self.Items._items:
            if hasattr(item, "Parent"):
                item.Parent = self

    def Delete(self) -> None:
        self.deleted = True


class FakeStore:
    def __init__(self, root_folder: FakeFolder) -> None:
        self._root_folder = root_folder

    def GetRootFolder(self) -> FakeFolder:
        return self._root_folder


class FakeAccountItem:
    Class = ItemType.ACCOUNT

    def __init__(self, name: str, address: str, root_folder: FakeFolder) -> None:
        self.DisplayName = name
        self.SmtpAddress = address
        self.DeliveryStore = FakeStore(root_folder)


class FakeNamespace:
    Class = ItemType.NAMESPACE

    def __init__(self, folder_map: dict[FolderType, FakeFolder]) -> None:
        self._folder_map = folder_map
        self.calls: list[FolderType] = []

    def GetDefaultFolder(self, folder_type: FolderType) -> FakeFolder:
        self.calls.append(folder_type)
        return self._folder_map[folder_type]


class ModelNavigationTests(unittest.TestCase):
    def test_account_matches_and_finds_nested_folder(self) -> None:
        reports = FakeFolder("Reports")
        inbox = FakeFolder("Inbox", subfolders=[reports])
        root = FakeFolder("Root", subfolders=[inbox])
        account = Account(FakeAccountItem("Primary Mailbox", "user@example.com", root))

        self.assertTrue(account.matches("Primary Mailbox"))
        self.assertTrue(account.matches("user@example.com"))
        self.assertFalse(account.matches("other@example.com"))
        self.assertEqual(account.find_folder("Inbox/Reports").name, "Reports")

    def test_folder_list_messages_filters_and_limits_results(self) -> None:
        first = FakeMailItem("First", unread=False)
        second = FakeMailItem("Second", unread=True)
        folder = Folder(FakeFolder("Inbox", items=[first, second]))

        unread_messages = folder.list_messages(unread_only=True)
        limited_messages = folder.list_messages(limit=1)

        self.assertEqual([message.subject for message in unread_messages], ["Second"])
        self.assertEqual([message.subject for message in limited_messages], ["First"])
        self.assertEqual(limited_messages[0].parent_folder.name, "Inbox")

    def test_folder_walk_returns_tree_entries(self) -> None:
        child = FakeFolder("Reports")
        folder = Folder(FakeFolder("Inbox", subfolders=[child]))

        entries = folder.walk(recursive=True, max_depth=1)

        self.assertEqual(
            [entry.as_row() for entry in entries],
            [("Inbox", 0, 1), ("Inbox/Reports", 1, 0)],
        )

    def test_default_folders_caches_by_enum_and_name(self) -> None:
        inbox = FakeFolder("Inbox")
        namespace = FakeNamespace({FolderType.INBOX: inbox})
        default_folders = DefaultFolders(namespace)

        first = default_folders.get(FolderType.INBOX)
        second = default_folders.get("inbox")

        self.assertIsNotNone(first)
        self.assertIs(first, second)
        self.assertEqual(namespace.calls, [FolderType.INBOX])


if __name__ == "__main__":
    unittest.main()
