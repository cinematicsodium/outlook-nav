from __future__ import annotations

from datetime import datetime

import typer
from tabulate import tabulate


def echo_table(rows: list[tuple[object, ...]], headers: list[str]) -> None:
    if not rows:
        typer.echo("No results.")
        return
    typer.echo(tabulate(rows, headers=headers, tablefmt="github"))


def format_datetime(value: datetime | None) -> str:
    if value is None:
        return ""
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value)


def truncate(value: str, width: int = 80) -> str:
    cleaned = " ".join(value.split())
    if len(cleaned) <= width:
        return cleaned
    return f"{cleaned[: width - 1].rstrip()}…"
