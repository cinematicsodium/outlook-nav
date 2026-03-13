from __future__ import annotations

import logging
from types import TracebackType

from ..enums import FolderType
from ..exceptions import OutlookConnectionError
from ..models.account import Account
from ..models.default_folders import DefaultFolders
from ..models.folder import Folder
from ..models.mail_item import MailItem
from ..protocols import OlApplication, OlDispatch, OlMailItem, OlNamespace
from ..utils import unpack_collection
from ..validation import validate_email

logger = logging.getLogger(__name__)


def _load_win32_client() -> OlDispatch:
    """Lazily import pywin32 to avoid import-time crashes on non-Windows systems."""
    try:
        import win32com.client as com_client  # type: ignore

        return com_client
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
            com_client = _load_win32_client()
            connection = com_client.Dispatch("Outlook.Application")
            namespace = connection.GetNamespace("MAPI")
            return cls(connection=connection, namespace=namespace)
        except Exception:
            logger.error("Error connecting to Outlook")
            raise

    def _ensure_connection(self) -> None:
        """Ensure Outlook connection and namespace are established."""
        try:
            if self._connection and self._namespace:
                return
            com_client = _load_win32_client()
            self._connection = com_client.Dispatch("Outlook.Application")
            self._namespace = self._connection.GetNamespace("MAPI")
        except Exception:
            logger.error("Error establishing Outlook connection")
            raise

    def _establish_user_account(self) -> None:
        """Set up the mailbox account based on the email address."""
        try:
            ol_accounts = unpack_collection(self._namespace.Accounts)
            accounts = [Account.from_outlook_item(account) for account in ol_accounts]
            valid_accts = [acct for acct in accounts if acct and acct]
            if len(accounts) == 1:
                self.mailbox_account = Account.from_outlook_item(accounts[0])
                return

            email = self.mailbox_address

            if not email:
                return

            addresses: list[str] = []

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
        """Return DefaultFolders instance for Outlook's default folders."""
        if self._default_folders is None:
            self._default_folders = DefaultFolders(self._require_namespace())
        return self._default_folders

    def get_folder(self, target_folder: str | FolderType) -> Folder | None:
        """Get a folder by name or default folder enum."""
        try:
            if isinstance(target_folder, str):
                return self.get_folder_by_name(target_folder)

            if isinstance(target_folder, FolderType):
                return self.get_default_folder(target_folder)

            raise ValueError("target_folder must be a string or FolderEnum")
        except Exception:
            logger.error("Error retrieving folder '%s'", target_folder)
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

            for folder in self.mailbox_account.folders:
                if str(folder.name).lower() == folder_name.lower():
                    return folder
            return None
        except Exception:
            logger.error("Error retrieving folder by name '%s'", folder_name)
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
                raise ValueError("folder_enum must be an instance of FolderEnum")
            return self.default_folders._get_default_folder(folder_enum)
        except Exception:
            logger.error("Error retrieving default folder '%s'", folder_enum)
            return None

    def list_mailboxes(self) -> list[Account]:
        """List all mailboxes/accounts."""
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
        """Get an account/mailbox by name."""
        if not isinstance(mailbox_name, str):
            raise ValueError("mailbox_name must be a string")
        target = mailbox_name.lower()
        for mailbox in self.list_mailboxes():
            if mailbox.name.lower() == target:
                return mailbox
        return None

    def create_email(self) -> MailItem | None:
        """Create a new email item."""
        try:
            mail_item: OlMailItem = self._require_connection().CreateItem(0)
            if self.mailbox_account:
                mail_item.SendUsingAccount = self.mailbox_account.ol_account_item
            return MailItem.from_outlook_item(mail_item)
        except Exception:
            logger.error("Error creating email")
            return None

    def list_emails(self, folder: Folder) -> list[MailItem]:
        """List emails in a folder."""
        if not isinstance(folder, Folder):
            raise ValueError("folder must be a Folder")

        items = folder.mail_items
        if isinstance(items, list):
            return items
        return []

    def move_email(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        """Move an email to another folder."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        return mail_item.move_to(destination)

    def send_email(self, mail_item: MailItem) -> None:
        """Send an email item."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.send()

    def delete_email(self, mail_item: MailItem) -> None:
        """Delete an email item."""
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")
        mail_item.delete()

    def create_folder(self, parent: Folder, folder_name: str) -> Folder | None:
        """Create a subfolder under a parent folder."""
        if not isinstance(parent, Folder):
            raise ValueError("parent must be a Folder")
        return parent.create_subfolder(folder_name)

    def delete_folder(self, folder: Folder) -> None:
        """Delete a folder."""
        if not isinstance(folder, Folder):
            raise ValueError("folder must be a Folder")
        folder.delete()

    def move_folder_item(
        self, source: Folder, mail_item: MailItem, destination: Folder
    ) -> MailItem | None:
        """Move a mail item from one folder to another."""
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
