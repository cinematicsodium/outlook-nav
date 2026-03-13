from __future__ import annotations

import typer

from ..context import create_client, get_state
from ..rendering import echo_table

mailboxes_app = typer.Typer(
    help="Inspect Outlook mailboxes that are visible to the current Outlook profile.",
    no_args_is_help=True,
)


@mailboxes_app.command("list")
def list_mailboxes(ctx: typer.Context) -> None:
    """List available Outlook mailboxes."""
    client = create_client(ctx)
    selected_mailbox = (get_state(ctx).mailbox or "").lower()
    default_account = client.mailbox_account

    rows: list[tuple[object, ...]] = []
    for account in client.list_mailboxes():
        matches_global_selection = selected_mailbox in {
            account.name.lower(),
            account.address.lower(),
        }
        is_default_account = (
            default_account is not None
            and account.name.lower() == default_account.name.lower()
            and account.address.lower() == default_account.address.lower()
        )
        rows.append(
            (
                account.name,
                account.address,
                len(account.folders),
                "yes" if matches_global_selection or is_default_account else "",
            )
        )

    echo_table(rows, headers=["Name", "Address", "Folders", "Selected"])
