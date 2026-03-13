from __future__ import annotations

from dataclasses import dataclass

import typer

from ..enums import FolderType
from ..exceptions import OutlookConnectionError
from ..models.account import Account
from ..models.folder import Folder
from ..models.outlook import OutlookApp


@dataclass(slots=True)
class CLIState:
    mailbox: str | None = None
    verbose: bool = False


DEFAULT_FOLDERS = {
    "archived mail": FolderType.ARCHIVED_MAIL,
    "archive": FolderType.ARCHIVED_MAIL,
    "calendar": FolderType.CALENDAR,
    "conflicts": FolderType.CONFLICTS,
    "contacts": FolderType.CONTACTS,
    "deleted items": FolderType.DELETED_ITEMS,
    "drafts": FolderType.DRAFTS,
    "inbox": FolderType.INBOX,
    "journal": FolderType.JOURNAL,
    "junk": FolderType.JUNK,
    "local failures": FolderType.LOCAL_FAILURES,
    "notes": FolderType.NOTES,
    "outbox": FolderType.OUTBOX,
    "sent mail": FolderType.SENT_MAIL,
}


def abort(message: str, exit_code: int = 1) -> None:
    typer.secho(message, fg=typer.colors.RED, err=True)
    raise typer.Exit(exit_code)


def get_state(ctx: typer.Context) -> CLIState:
    state = ctx.obj
    if isinstance(state, CLIState):
        return state
    return CLIState()


def create_client(ctx: typer.Context) -> OutlookApp:
    state = get_state(ctx)
    mailbox_address = state.mailbox if state.mailbox and "@" in state.mailbox else None

    try:
        return OutlookApp(mailbox_address=mailbox_address)
    except OutlookConnectionError as exc:
        abort(
            f"{exc} This CLI requires Microsoft Outlook and pywin32 on Windows."
        )
    except Exception as exc:
        abort(f"Unable to connect to Outlook: {exc}")


def resolve_account(
    client: OutlookApp,
    ctx: typer.Context,
    mailbox: str | None = None,
) -> Account:
    state = get_state(ctx)
    selected_mailbox = mailbox or state.mailbox

    if selected_mailbox:
        if client.mailbox_account is not None:
            names = {
                client.mailbox_account.name.lower(),
                client.mailbox_account.address.lower(),
            }
            if selected_mailbox.lower() in names:
                return client.mailbox_account

        account = client.get_mailbox(selected_mailbox)
        if account is not None:
            return account

        abort(f"Mailbox not found: {selected_mailbox}")

    if client.mailbox_account is not None:
        return client.mailbox_account

    mailboxes = client.list_mailboxes()
    if len(mailboxes) == 1:
        return mailboxes[0]

    abort("Multiple mailboxes are available. Use --mailbox to select one.")


def resolve_folder(
    client: OutlookApp,
    ctx: typer.Context,
    folder_path: str,
    mailbox: str | None = None,
) -> Folder:
    normalized_path = folder_path.strip()
    if not normalized_path:
        abort("Folder path cannot be empty.")

    selected_mailbox = mailbox or get_state(ctx).mailbox
    account = resolve_account(client, ctx, mailbox=mailbox)
    folder = account.find_folder(normalized_path)
    if folder is not None:
        return folder

    folder_enum = DEFAULT_FOLDERS.get(normalized_path.lower())
    if folder_enum is not None and not selected_mailbox:
        default_folder = client.get_default_folder(folder_enum)
        if default_folder is not None:
            return default_folder

    abort(
        f"Folder not found in mailbox '{account.name}': {normalized_path}"
    )
