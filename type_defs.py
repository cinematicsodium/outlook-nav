from pathlib import Path
from typing import TYPE_CHECKING, Literal, TypeAlias, TypeVar

if TYPE_CHECKING:
    from .models.node import ItemModel

LowerStr: TypeAlias = str
StrPath: TypeAlias = str | Path

T = TypeVar("T", covariant=True)

ModelT = TypeVar("ModelT", bound="ItemModel")

RawT = TypeVar("RawT")

TableFmtStr = Literal[
    "simple",
    "plain",
    "grid",
    "simple_grid",
    "rounded_grid",
    "heavy_grid",
    "mixed_grid",
    "double_grid",
    "fancy_grid",
    "outline",
    "simple_outline",
    "rounded_outline",
    "heavy_outline",
    "mixed_outline",
    "double_outline",
    "fancy_outline",
    "github",
    "pipe",
    "orgtbl",
    "jira",
    "presto",
    "pretty",
    "psql",
    "rst",
    "mediawiki",
    "moinmoin",
    "youtrack",
    "html",
    "unsafehtml",
    "latex",
    "latex_raw",
    "latex_booktabs",
    "latex_longtable",
    "tsv",
    "textile",
    "asciidoc",
]
