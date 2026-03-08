from __future__ import annotations

import logging
from types import TracebackType
from typing import TypeAlias, cast

from outlook.constants import DefaultFolderEnum
from outlook.models.account import Account
from outlook.models.default_folders import DefaultFolders
from outlook.models.folder import Folder
from outlook.models.mail_item import MailItem
from outlook.protocols import OlApplication, OlDispath, OlMailItem, OlNamespace
from outlook.utils import validate_email

logger = logging.getLogger(__name__)

OlApplication: TypeAlias = OlApplication | None
OlNamespace: TypeAlias = OlNamespace | None


class OutlookApp:
    def __init__(
        self,
        mailbox_address: str | None = None,
        connection: OlApplication = None,
        namespace: OlNamespace = None,
    ):
        self.mailbox_address = validate_email(mailbox_address)
        self._connection = connection
        self._namespace = namespace

        self._default_folders: DefaultFolders | None = None
        self.mailbox_account: Account | None = None

        self._establish_user_account()

    @classmethod
    def connect(cls):
        try:
            com_client = _load_win32_client()
            connection = com_client.Dispatch("Outlook.Application")
            namespace = connection.GetNamespace("MAPI")
            return cls(connection=connection, namespace=namespace)
        except Exception:
            logger.error("Error connecting to Outlook")
            raise

    def _establish_user_account(self) -> None:
        try:
            if not (count := self._namespace.Accounts.Count):
                return

            accounts = [self._namespace.Accounts.Item(i) for i in range(1, count + 1)]
            if len(accounts) == 1:
                self.mailbox_account = Account.from_outlook_item(accounts[0])
                return

            email = self.mailbox_address

            if not email:
                return

            addresses = []

            for account in accounts:
                acct_address = account.SmtpAddress.lower()
                addresses.append(acct_address)

                if acct_address == email:
                    self.mailbox_account = Account.from_outlook_item(account)
                    logger.info(f"Successfully set default account to: {email}")
                    return

            logger.warning(f"Account {email} not found. Available accounts are:")
            logger.warning("\n".join(f"- {address}" for address in addresses))

        except Exception:
            return

    @property
    def default_folders(self) -> DefaultFolders:
        """Returns an instance of DefaultFolders for easy access to Outlook's default folders."""
        if self._default_folders is None:
            self._default_folders = DefaultFolders(self._require_namespace())
        return self._default_folders

    def get_folder(self, target_folder: str | DefaultFolderEnum) -> Folder | None:
        try:
            if isinstance(target_folder, str):
                return self.get_folder_by_name(target_folder)

            if isinstance(target_folder, DefaultFolderEnum):
                return self.get_default_folder(target_folder)

            raise ValueError("target_folder must be a string or FolderEnum")
        except Exception:
            logger.error("Error retrieving folder '%s'", target_folder)
            return None

    def get_folder_by_name(self, folder_name: str) -> Folder | None:
        try:

            if not self.mailbox_account:
                logger.error(
                    "'OutlookApp.mailbox_address' must be defined "
                    "to perform a folder search."
                )
                return
            if not isinstance(folder_name, str):
                raise ValueError("folder_name must be a string")

            for folder in self.mailbox_account.folders:
                if str(folder.name).lower() == folder_name.lower():
                    return folder
            return None
        except Exception:
            logger.error("Error retrieving folder by name '%s'", folder_name)
            return None

    def get_mailbox_folder(self, mailbox_name: str, folder_name: str) -> Folder | None:
        mailbox = self.get_mailbox(mailbox_name)
        if mailbox is None:
            return None
        return mailbox.get_folder(folder_name)

    def get_default_folder(self, folder_enum: DefaultFolderEnum) -> Folder | None:
        try:
            if not isinstance(folder_enum, DefaultFolderEnum):
                raise ValueError("folder_enum must be an instance of FolderEnum")
            return self.default_folders.resolve(folder_enum)
        except Exception:
            logger.error("Error retrieving default folder '%s'", folder_enum)
            return None

    def list_mailboxes(self) -> list[Account]:
        accounts: list[Account] = []
        try:
            for mailbox in self._require_namespace().Folders:
                account = Account.from_outlook_item(mailbox)
                if account:
                    accounts.append(account)
        except Exception:
            logger.error("Error listing mailboxes")
        return accounts

    def get_mailbox(self, mailbox_name: str) -> Account | None:
        if not isinstance(mailbox_name, str):
            raise ValueError("mailbox_name must be a string")
        target = mailbox_name.lower()
        for mailbox in self.list_mailboxes():
            if mailbox.name.lower() == target:
                return mailbox
        return None

    def create_email(self) -> MailItem | None:
        try:
            mail_item: OlMailItem = self._require_connection().CreateItem(0)
            if self.mailbox_account:
                mail_item.SendUsingAccount = self.mailbox_account.ol_account_item
            return MailItem.from_outlook_item(mail_item)
        except Exception:
            logger.error("Error creating email")
            return None

    def list_emails(self, folder: Folder) -> list[MailItem]:
        if not isinstance(folder, Folder):
            raise ValueError("folder must be a Folder")

        items = folder.mail_items
        if isinstance(items, list):
            return items
        return []

    def move_email(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        return mail_item.move_to(destination)

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
            if self._connection is not None:
                self._connection.Quit()
        except Exception:
            logger.error("Error closing Outlook connection")
        finally:
            self._connection = None
            self._namespace = None
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

    def _require_connection(self) -> OlApplication:
        if self._connection is None:
            raise RuntimeError(
                "Outlook connection is not available. Call connect() first."
            )
        return self._connection

    def _require_namespace(self) -> OlNamespace:
        if self._namespace is None:
            raise RuntimeError(
                "Outlook namespace is not available. Call connect() first."
            )
        return self._namespace


def _load_win32_client() -> OlDispath:
    """
    Lazily import pywin32 to avoid import-time crashes on non-Windows systems.
    """
    try:
        import win32com.client as com_client  # type: ignore

        return cast(OlDispath, com_client)
    except ImportError as exc:
        raise RuntimeError(
            "win32com.client is required to use OutlookApp. Install pywin32 on Windows."
        ) from exc
