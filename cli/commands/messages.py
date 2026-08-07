from __future__ import annotations
import typer
from ..context import create_client, resolve_folder
from ..rendering import echo_table, format_datetime, truncate

app = typer.Typer(
    help="Inspect messages in an Outlook folder.",
    no_args_is_help=True,
)


@app.command("list")
def list_messages(
    ctx: typer.Context,
    folder_path: str = typer.Argument(
        ...,
        help="Folder name or slash-delimited path, such as 'Inbox' or 'Inbox/Reports'.",
    ),
    account: str | None = typer.Option(
        None,
        "--account",
        help="Mailbox display name or SMTP address to inspect for this command.",
    ),
    limit: int = typer.Option(
        20,
        "--limit",
        "-n",
        min=1,
        max=500,
        help="Maximum number of messages to display.",
    ),
    unread_only: bool = typer.Option(
        False,
        "--unread-only",
        help="Show only unread messages.",
    ),
    include_body_preview: bool = typer.Option(
        False,
        "--include-body-preview",
        help="Include a short plain-text preview column in the output.",
    ),
) -> None:
    """List recent messages from a folder."""
    client = create_client(ctx)
    folder = resolve_folder(client, ctx, folder_path, account=account)
    messages = folder.list_messages(limit=limit, unread_only=unread_only)
    rows: list[tuple[object, ...]] = []
    for message in messages:
        received = format_datetime(message.received_at or message.sent_at)
        subject = truncate(message.subject or "(no subject)", width=60)
        sender = truncate(message.sender_address or message.sender_name, width=36)
        unread = "yes" if message.unread else ""
        if include_body_preview:
            preview = truncate(message.body, width=70)
            rows.append((received, unread, sender, subject, preview))
            continue
        rows.append((received, unread, sender, subject))
    headers = ["Received", "Unread", "From", "Subject"]
    if include_body_preview:
        headers.append("Preview")
    echo_table(rows, headers=headers)
