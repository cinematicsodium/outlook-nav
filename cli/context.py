from __future__ import annotations

from dataclasses import dataclass
from typing import NoReturn

import typer

from ..enums import FolderEnum
from ..exceptions import OutlookError
from ..models.account import Account
from ..models.folder import Folder
from ..models.outlook import Outlook


@dataclass(slots=True)
class CLIState:
    """Global account selection and logging options for a CLI invocation."""

    account: str | None = None
    verbose: bool = False


DEFAULT_FOLDERS = {
    "archived mail": FolderEnum.ARCHIVED_MAIL,
    "archive": FolderEnum.ARCHIVED_MAIL,
    "calendar": FolderEnum.CALENDAR,
    "conflicts": FolderEnum.CONFLICTS,
    "contacts": FolderEnum.CONTACTS,
    "deleted items": FolderEnum.DELETED_ITEMS,
    "drafts": FolderEnum.DRAFTS,
    "inbox": FolderEnum.INBOX,
    "journal": FolderEnum.JOURNAL,
    "junk": FolderEnum.JUNK,
    "local failures": FolderEnum.LOCAL_FAILURES,
    "notes": FolderEnum.NOTES,
    "outbox": FolderEnum.OUTBOX,
    "sent mail": FolderEnum.SENT_MAIL,
}


def abort(message: str, exit_code: int = 1) -> NoReturn:
    """Print an error and terminate the CLI command.

    Parameters
    ----------
    message : str
        Error message to print.
    exit_code : int, default=1
        Process exit status.
    """
    typer.secho(message, fg=typer.colors.RED, err=True)
    raise typer.Exit(exit_code)


def get_state(ctx: typer.Context) -> CLIState:
    """Return the CLI state stored on a Typer context.

    Parameters
    ----------
    ctx : typer.Context
        Current command context.

    Returns
    -------
    CLIState
        Stored state, or an empty default state.
    """
    state = ctx.obj
    if isinstance(state, CLIState):
        return state
    return CLIState()


def create_client(ctx: typer.Context) -> Outlook:
    """Create an Outlook client from the current CLI state.

    Parameters
    ----------
    ctx : typer.Context
        Current command context.

    Returns
    -------
    Outlook
        Connected Outlook client.
    """
    state = get_state(ctx)
    try:
        return Outlook(address=state.account)
    except OutlookError as exc:
        abort(str(exc))
    except Exception as exc:  # ruff: ignore[BLE001]
        abort(f"Unable to connect to Outlook: {exc}")


def resolve_account(
    client: Outlook,
    ctx: typer.Context,
    account: str | None = None,
) -> Account:
    """Resolve the account selected for a command.

    Parameters
    ----------
    client : Outlook
        Connected Outlook client.
    ctx : typer.Context
        Current command context.
    account : str, optional
        Command-specific display name or SMTP address.

    Returns
    -------
    Account
        Selected account.
    """
    state = get_state(ctx)
    selection = account or state.account
    if selection:
        if client.account is not None:  # ruff: ignore[SIM102]
            if client.account.matches(selection):
                return client.account
        resolved = client.find_account(selection)
        if resolved is not None:
            return resolved
        abort(f"Account not found: {selection}")
    if client.account is not None:
        return client.account
    accounts = client.accounts
    if len(accounts) == 1:
        return accounts[0]
    abort("Multiple accounts are available. Use --account to select one.")


def resolve_folder(
    client: Outlook,
    ctx: typer.Context,
    folder_path: str,
    account: str | None = None,
) -> Folder:
    """Resolve a folder path for a CLI command.

    Parameters
    ----------
    client : Outlook
        Connected Outlook client.
    ctx : typer.Context
        Current command context.
    folder_path : str
        Folder name or slash-delimited path.
    account : str, optional
        Command-specific display name or SMTP address.

    Returns
    -------
    Folder
        Matching folder.
    """
    normalized_path = folder_path.strip()
    if not normalized_path:
        abort("Folder path cannot be empty.")
    resolved_account = resolve_account(client, ctx, account=account)
    folder = resolved_account.find_folder(normalized_path)
    if folder is not None:
        return folder
    folder_enum = DEFAULT_FOLDERS.get(normalized_path.lower())
    if folder_enum is not None:
        default_folder = resolved_account.default_folder(folder_enum)
        if default_folder is not None:
            return default_folder
    abort(f"Folder not found in account '{resolved_account.name}': {normalized_path}")
