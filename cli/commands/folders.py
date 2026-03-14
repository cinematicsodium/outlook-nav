from __future__ import annotations

import typer

from ..context import create_client, resolve_account, resolve_folder
from ..rendering import echo_table

app = typer.Typer(
    help="Browse mailbox folders and inspect mailbox structure.",
    no_args_is_help=True,
)


@app.command("list")
def list_folders(
    ctx: typer.Context,
    mailbox: str | None = typer.Option(
        None,
        "--mailbox",
        help="Mailbox display name or SMTP address to inspect for this command.",
    ),
    root: str | None = typer.Option(
        None,
        "--root",
        help="Folder path to use as the starting point, such as 'Inbox/Reports'.",
    ),
    recursive: bool = typer.Option(
        False,
        "--recursive",
        "-r",
        help="Walk child folders recursively.",
    ),
    max_depth: int = typer.Option(
        2,
        "--max-depth",
        min=1,
        help="Maximum child depth to include when using --recursive.",
    ),
) -> None:
    """List folders for a mailbox or a specific folder branch."""
    client = create_client(ctx)

    if root:
        base_folder = resolve_folder(client, ctx, root, mailbox=mailbox)
        rows = [
            entry.as_row()
            for entry in base_folder.walk(
                recursive=recursive,
                max_depth=max_depth,
            )
        ]
    else:
        account = resolve_account(client, ctx, mailbox=mailbox)
        rows: list[tuple[object, ...]] = []
        for folder in account.folders:
            rows.extend(
                entry.as_row()
                for entry in folder.walk(
                    recursive=recursive,
                    max_depth=max_depth,
                )
            )

    echo_table(rows, headers=["Folder", "Depth", "Subfolders"])
