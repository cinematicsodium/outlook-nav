from __future__ import annotations
import logging
import typer
from .commands import drafts, folders, accounts, messages
from .context import CLIState

app = typer.Typer(
    help="Command-line tools for inspecting Outlook accounts and managing messages.",
    no_args_is_help=True,
    add_completion=False,
)
app.add_typer(accounts.app, name="accounts")
app.add_typer(folders.app, name="folders")
app.add_typer(messages.app, name="messages")
app.add_typer(drafts.app, name="drafts")


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.getLogger().setLevel(level)
    logging.getLogger("rich").setLevel(level)


@app.callback()
def main(
    ctx: typer.Context,
    account: str | None = typer.Option(
        None,
        "--account",
        "-m",
        help="Default account display name or SMTP address for account-aware commands.",
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
    ctx.obj = CLIState(account=account, verbose=verbose)


def run() -> None:
    app()
