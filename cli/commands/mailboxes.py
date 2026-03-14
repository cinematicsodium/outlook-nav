from __future__ import annotations

import typer

from ..context import create_client, get_state
from ..rendering import echo_table

app = typer.Typer(
    help="Inspect Outlook mailboxes that are visible to the current Outlook profile.",
    no_args_is_help=True,
)


@app.command("list")
def list_mailboxes(ctx: typer.Context) -> None:
    """List available Outlook mailboxes."""
    client = create_client(ctx)
    selected_mailbox = get_state(ctx).mailbox
    default_account = client.mailbox_account

    rows: list[tuple[object, ...]] = []
    for account in client.list_mailboxes():
        matches_global_selection = account.matches(selected_mailbox)
        is_default_account = default_account is not None and account.matches(
            default_account.address
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
