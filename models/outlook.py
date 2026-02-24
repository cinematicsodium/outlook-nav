from __future__ import annotations

import logging
from types import TracebackType
from typing import cast

from constants import FolderEnum
from models.account import MailboxAccount
from models.default_folders import DefaultFolders
from models.folder import Folder
from models.mail_item import MailItem
from protocols import (
    DispatchModuleProtocol,
    NamespaceProtocol,
    OutlookApplicationProtocol,
)
from utils import validate_email

logger = logging.getLogger(__name__)


class OutlookApp:
    def __init__(self, mailbox_address: str | None = None):
        self.connection: OutlookApplicationProtocol | None = None
        self.namespace: NamespaceProtocol | None = None
        self.mailbox_address: str | None = None
        self._default_folders: DefaultFolders | None = None
        if mailbox_address is not None:
            self.mailbox_address = validate_email(mailbox_address)
        self.connect()

    def connect(self) -> None:
        com_client = _load_win32_client()
        try:
            self.connection = com_client.Dispatch("Outlook.Application")
            self.namespace = self.connection.GetNamespace("MAPI")
            self._default_folders = None
        except Exception:
            logger.exception("Error connecting to Outlook")
            raise

    @property
    def default_folders(self) -> DefaultFolders:
        """Returns an instance of DefaultFolders for easy access to Outlook's default folders."""
        if self._default_folders is None:
            self._default_folders = DefaultFolders(self._require_namespace())
        return self._default_folders

    def get_folder(self, target_folder: str | FolderEnum) -> Folder | None:
        try:
            if isinstance(target_folder, str):
                return self.get_folder_by_name(target_folder)

            if isinstance(target_folder, FolderEnum):
                return self.get_default_folder(target_folder)

            raise ValueError("target_folder must be a string or FolderEnum")
        except Exception:
            logger.exception("Error retrieving folder '%s'", target_folder)
            return None

    def get_folder_by_name(self, folder_name: str) -> Folder | None:
        try:
            if not isinstance(folder_name, str):
                raise ValueError("folder_name must be a string")
            folders = self._require_namespace().Folders
            for folder in folders:
                if str(folder.Name).lower() == folder_name.lower():
                    return Folder(folder)
            return None
        except Exception:
            logger.exception("Error retrieving folder by name '%s'", folder_name)
            return None

    def list_mailboxes(self) -> list[MailboxAccount]:
        accounts: list[MailboxAccount] = []
        try:
            for mailbox in self._require_namespace().Folders:
                accounts.append(MailboxAccount(mailbox))
        except Exception:
            logger.exception("Error listing mailboxes")
        return accounts

    def get_mailbox(self, mailbox_name: str) -> MailboxAccount | None:
        if not isinstance(mailbox_name, str):
            raise ValueError("mailbox_name must be a string")
        target = mailbox_name.lower()
        for mailbox in self.list_mailboxes():
            if mailbox.name.lower() == target:
                return mailbox
        return None

    def get_mailbox_folder(self, mailbox_name: str, folder_name: str) -> Folder | None:
        mailbox = self.get_mailbox(mailbox_name)
        if mailbox is None:
            return None
        return mailbox.get_folder(folder_name)

    def get_default_folder(self, folder_enum: FolderEnum) -> Folder | None:
        try:
            if not isinstance(folder_enum, FolderEnum):
                raise ValueError("folder_enum must be an instance of FolderEnum")
            return self.default_folders.resolve(folder_enum)
        except Exception:
            logger.exception("Error retrieving default folder '%s'", folder_enum)
            return None

    def create_email(self) -> MailItem | None:
        try:
            mail_item = self._require_connection().CreateItem(0)
            return MailItem(mail_item)
        except Exception:
            logger.exception("Error creating email")
            return None

    def list_emails(self, folder: Folder) -> list[MailItem]:
        if not isinstance(folder, Folder):
            raise ValueError("folder must be a Folder")

        items = folder.items
        if isinstance(items, list):
            return items
        return []

    def move_email(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        return mail_item.move(destination)

    def send_email(self, mail_item: MailItem) -> None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.send()

    def delete_email(self, mail_item: MailItem) -> None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def create_folder(self, parent: Folder, folder_name: str) -> Folder | None:
        if not isinstance(parent, Folder):
            raise ValueError("parent must be a Folder")
        return parent.create_subfolder(folder_name)

    def delete_folder(self, folder: Folder) -> None:
        if not isinstance(folder, Folder):
            raise ValueError("folder must be a Folder")
        folder.delete()

    def move_folder_item(
        self, source: Folder, mail_item: MailItem, destination: Folder
    ) -> MailItem | None:
        if not isinstance(source, Folder):
            raise ValueError("source must be a Folder")
        return source.move_item(mail_item, destination)

    def close(self) -> None:
        """Release COM resources to avoid leaks."""
        try:
            if self.connection is not None:
                self.connection.Quit()
        except Exception:
            logger.exception("Error closing Outlook connection")
        finally:
            self.connection = None
            self.namespace = None
            self._default_folders = None

    def __enter__(self) -> OutlookApp:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()

    def _require_connection(self) -> OutlookApplicationProtocol:
        if self.connection is None:
            raise RuntimeError("Outlook connection is not available. Call connect() first.")
        return self.connection

    def _require_namespace(self) -> NamespaceProtocol:
        if self.namespace is None:
            raise RuntimeError("Outlook namespace is not available. Call connect() first.")
        return self.namespace


def _load_win32_client() -> DispatchModuleProtocol:
    """
    Lazily import pywin32 to avoid import-time crashes on non-Windows systems.
    """
    try:
        import win32com.client as com_client  # type: ignore

        return cast(DispatchModuleProtocol, com_client)
    except ImportError as exc:
        raise RuntimeError(
            "win32com.client is required to use OutlookApp. Install pywin32 on Windows."
        ) from exc
