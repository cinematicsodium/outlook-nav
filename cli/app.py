from __future__ import annotations

import logging

import typer

from .commands.drafts import drafts_app
from .commands.folders import folders_app
from .commands.mailboxes import mailboxes_app
from .commands.messages import messages_app
from .context import CLIState

app = typer.Typer(
    help="Command-line tools for inspecting Outlook mailboxes and managing messages.",
    no_args_is_help=True,
    add_completion=False,
)

app.add_typer(mailboxes_app, name="mailboxes")
app.add_typer(folders_app, name="folders")
app.add_typer(messages_app, name="messages")
app.add_typer(drafts_app, name="drafts")


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.getLogger().setLevel(level)
    logging.getLogger("rich").setLevel(level)


@app.callback()
def main(
    ctx: typer.Context,
    mailbox: str | None = typer.Option(
        None,
        "--mailbox",
        "-m",
        help="Default mailbox display name or SMTP address for mailbox-aware commands.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Enable verbose logging for Outlook operations.",
    ),
) -> None:
    """Configure global CLI options."""
    _configure_logging(verbose)
    ctx.obj = CLIState(mailbox=mailbox, verbose=verbose)


def run() -> None:
    app()
