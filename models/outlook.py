from __future__ import annotations

import logging
from types import TracebackType
from typing import cast

from ..enums import FolderType
from ..exceptions import OutlookConnectionError
from ..models.account import Account
from ..models.default_folders import DefaultFolders
from ..models.folder import Folder
from ..models.mail_item import MailItem
from ..protocols import OlApplication, OlDispatch, OlMailItem, OlNamespace
from ..type_defs import T
from ..utils import unpack_collection
from ..validation import validate_email

logger = logging.getLogger(__name__)


def _load_win32_client() -> OlDispatch:
    """Lazily import pywin32 to avoid import-time crashes on non-Windows systems."""
    try:
        from win32com import client  # type: ignore

        return client
    except ImportError as exc:
        raise OutlookConnectionError(
            "win32com.client is required to use OutlookApp. Install pywin32 on Windows."
        ) from exc


class OutlookApp:
    """Main Outlook application interface."""

    def __init__(
        self,
        mailbox_address: str | None = None,
        connection: OlApplication = None,
        namespace: OlNamespace = None,
    ) -> None:
        """Initialize OutlookApp instance."""
        self.mailbox_address: str | None = validate_email(mailbox_address)
        self._connection: OlApplication | None = connection
        self._namespace: OlNamespace | None = namespace

        self._default_folders: DefaultFolders | None = None
        self.mailbox_account: Account | None = None

        self._ensure_connection()
        self._establish_user_account()

    @classmethod
    def connect(cls) -> OutlookApp:
        """Connect to Outlook and return an OutlookApp instance."""
        try:
            client = _load_win32_client()
            connection = client.Dispatch("Outlook.Application")
            namespace = connection.GetNamespace("MAPI")

            return cls(connection=connection, namespace=namespace)
        except Exception:
            logger.exception("Error connecting to Outlook")
            raise

    @property
    def default_folders(self) -> DefaultFolders:
        """Return DefaultFolders instance for Outlook's default folders."""
        if self._default_folders is None:
            namespace = self._require_namespace()

            default_folders = DefaultFolders.from_outlook_item(namespace)
            if default_folders is None:
                logger.error(
                    "Unable to initialize default Outlook folders from namespace"
                )
                raise OutlookConnectionError(
                    "Unable to initialize default Outlook folders."
                )
            self._default_folders = default_folders
        return self._default_folders

    def get_folder(self, target_folder: str | FolderType) -> Folder | None:
        """Get a folder by name or default folder type identifier."""
        try:
            if isinstance(target_folder, str):
                return self.get_folder_by_name(target_folder)

            if isinstance(target_folder, FolderType):
                return self.get_default_folder(target_folder)

            raise ValueError("target_folder must be a string or FolderType enum")
        except Exception:
            logger.exception("Error retrieving folder '%s'", target_folder)
            return None

    def get_folder_by_name(self, folder_name: str) -> Folder | None:
        """Get a folder by its name."""
        try:
            if not self.mailbox_account:
                logger.error(
                    "'OutlookApp.mailbox_address' must be defined "
                    "to perform a folder search."
                )
                return
            if not isinstance(folder_name, str):
                raise ValueError("folder_name must be a string")
            return self.mailbox_account.get_folder(folder_name)
        except Exception:
            logger.exception("Error retrieving folder by name '%s'", folder_name)
            return None

    def get_mailbox_folder(self, mailbox_name: str, folder_name: str) -> Folder | None:
        """Get a folder from a specific mailbox."""
        mailbox = self.get_mailbox(mailbox_name)
        if mailbox is None:
            return None
        return mailbox.get_folder(folder_name)

    def get_default_folder(self, folder_enum: FolderType) -> Folder | None:
        """Get a default folder by enum."""
        try:
            if not isinstance(folder_enum, FolderType):
                raise ValueError("folder_enum must be an instance of FolderType")
            return self.default_folders.get(folder_enum)
        except Exception:
            logger.exception("Error retrieving default folder '%s'", folder_enum)
            return None

    def list_mailboxes(self) -> list[Account]:
        """List all mailboxes/accounts."""
        try:
            accounts = self._list_accounts()
            return accounts
        except Exception:
            logger.exception("Error listing mailboxes")
            return []

    def get_mailbox(self, mailbox_name: str) -> Account | None:
        """Get an account/mailbox by name."""
        mailbox_name = self._ensure_item_type(mailbox_name, str)
        for mailbox in self.list_mailboxes():
            if mailbox.matches(mailbox_name):
                return mailbox
        return None

    def create_email(self) -> MailItem | None:
        """Create a new email item."""
        try:
            connection = self._require_connection()
            mail_item: OlMailItem = connection.CreateItem(0)
            if self.mailbox_account:
                mail_item.SendUsingAccount = self.mailbox_account._ol_account_item
            return MailItem.from_outlook_item(mail_item)
        except Exception:
            logger.exception("Error creating email")
            return None

    def list_emails(self, folder: Folder) -> list[MailItem]:
        """List emails in a folder."""
        folder = self._ensure_item_type(folder, Folder)
        return folder.list_messages()

    def move_email(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        """Move an email to another folder."""
        mail_item = self._ensure_item_type(mail_item, MailItem)
        destination = self._ensure_item_type(destination, Folder)
        return mail_item.move_to(destination)

    def send_email(self, mail_item: MailItem) -> None:
        """Send an email item."""
        mail_item = self._ensure_item_type(mail_item, MailItem)
        mail_item.send()

    def delete_email(self, mail_item: MailItem) -> None:
        """Delete an email item."""
        mail_item = self._ensure_item_type(mail_item, MailItem)
        mail_item.delete()

    def create_folder(self, parent: Folder, folder_name: str) -> Folder | None:
        """Create a subfolder under a parent folder."""
        parent = self._ensure_item_type(parent, Folder)
        folder_name = self._ensure_item_type(folder_name, str)

        return parent.create_subfolder(folder_name)

    def delete_folder(self, folder: Folder) -> None:
        """Delete a folder."""
        folder = self._ensure_item_type(folder, Folder)
        folder.delete()

    def move_folder(
        self, source: Folder, mail_item: MailItem, destination: Folder
    ) -> MailItem | None:
        """Move a mail item from one folder to another."""
        source = self._ensure_item_type(source, Folder)
        mail_item = self._ensure_item_type(mail_item, MailItem)
        destination = self._ensure_item_type(destination, Folder)
        return source.move_item(mail_item, destination)

    def close(self) -> None:
        """Release COM resources to avoid leaks."""
        try:
            if self._connection is not None:
                self._connection.Quit()
        except Exception:
            logger.exception("Error closing Outlook connection")
        finally:
            self._connection = None
            self._namespace = None
            self._default_folders = None

    def __enter__(self) -> OutlookApp:
        """Enter context manager."""
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        """Exit context manager and close resources."""
        self.close()

    def __repr__(self):
        account_email = self.mailbox_account.email_address
        return (
            f"OutlookApp("
            f"mailbox_address={self.mailbox_address!r}, "
            f"account={account_email!r}, "
            f"connected={self._connection is not None}, "
            f"namespace={self._namespace is not None}"
            f")"
        )

    def __str__(self) -> str:
        return self.__repr__()

    def _ensure_connection(self) -> None:
        """Ensure Outlook connection and namespace are established."""
        try:
            if self._connection and self._namespace:
                return

            client = _load_win32_client()
            self._connection = client.Dispatch("Outlook.Application")
            self._namespace = self._connection.GetNamespace("MAPI")

        except Exception:
            logger.exception("Error establishing Outlook connection")
            raise

    def _establish_user_account(self) -> None:
        """Set up the mailbox account based on the email address."""
        try:
            accounts = self._list_accounts()
            if len(accounts) == 1:
                self.mailbox_account = accounts[0]
                return

            email = self.mailbox_address
            if not email:
                return

            matched_account = self._find_account_by_email(accounts, email)
            if matched_account:
                self.mailbox_account = matched_account
                logger.info(f"Successfully set default account to: {email}")
                return

            self._log_account_not_found(email, accounts)
        except Exception:
            if self.mailbox_address:
                logger.warning(
                    "Unable to resolve Outlook mailbox account for '%s'",
                    self.mailbox_address,
                    exc_info=True,
                )
            else:
                logger.warning(
                    "Unable to inspect Outlook accounts during initialization",
                    exc_info=True,
                )
            return

    def _find_account_by_email(
        self, accounts: list[Account], email: str
    ) -> Account | None:
        """Find an account matching the given email address."""
        normalized_email = email.lower()
        for account in accounts:
            if account.matches(normalized_email):
                return account
        return None

    def _log_account_not_found(self, email: str, accounts: list[Account]) -> None:
        """Log a warning if the account is not found."""
        addresses = [account.email_address.lower() for account in accounts]
        warning_message = (
            f"Account '{email}' not found among {len(accounts)} accounts. "
        )
        warning_message += "Available accounts:\n"
        warning_message += "\n".join(f"- {address}" for address in addresses)
        logger.warning(warning_message)

    def _require_connection(self) -> OlApplication:
        """Return Outlook connection or raise if unavailable."""
        if self._connection is None:
            raise OutlookConnectionError(
                "Outlook connection is not available. Call connect() first."
            )
        return self._connection

    def _require_namespace(self) -> OlNamespace:
        """Return Outlook namespace or raise if unavailable."""
        if self._namespace is None:
            raise OutlookConnectionError(
                "Outlook namespace is not available. Call connect() first."
            )
        return self._namespace

    def _list_accounts(self, namespace: OlNamespace | None = None) -> list[Account]:
        namespace = namespace or self._require_namespace()
        accounts = unpack_collection(namespace.Accounts, transformer=Account)
        return accounts

    def _ensure_item_type(self, item: object, target_type: type[T]) -> T:
        """Helper to validate and return an Outlook item of the expected type."""
        if not isinstance(item, target_type):
            item_type_name = type(item).__name__
            target_type_name = target_type.__name__
            msg = f"Expected item of type {target_type_name}, got {item_type_name}"
            raise TypeError(msg)
        return cast(T, item)
