import logging

from rich.console import Console
from rich.logging import RichHandler
from rich.traceback import install

console = Console()

install(console=console, show_locals=True)

LOG_FILE = "app.log"


def _reduce_logs(max_len=10_000):
    try:
        from pathlib import Path

        path = Path(LOG_FILE).resolve()

        if not path.exists():
            return

        path_bytes = path.read_bytes()
        text = path_bytes.decode(errors="ignore")
        logs = [line.strip() for line in text.splitlines()]

        if len(logs) < max_len:
            return

        tail_lines_reversed = list(reversed(logs))[:max_len]
        trimmed_logs = list(reversed(tail_lines_reversed))
        text = "\n".join(trimmed_logs)
        path.write_text(text, encoding="utf-8")
    except Exception:
        return


# LEVEL = logging.DEBUG
LEVEL = logging.INFO


FORMATTER = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")


file_handler = logging.FileHandler(LOG_FILE)
file_handler.setLevel(LEVEL)
file_handler.setFormatter(FORMATTER)

rich_handler = RichHandler()
rich_handler.setLevel(LEVEL)

logging.basicConfig(
    level=LEVEL,
    format="%(message)s",
    datefmt="[%X]",
    handlers=[rich_handler, file_handler],
)

log = logging.getLogger("rich")
log.rule = console.rule

_reduce_logs()

if __name__ == "__main__":
    log.debug("This is a debug message.")
    log.info("This is an info message.")
    log.warning("This is a warning message.")
    log.error("This is an error message.")
    log.critical("This is a critical message.")
