"""Shapes of the dict-valued options.

Most of this library's surface is a dict handed to ``sheet()``, and until these
existed a type checker saw ``Dict[str, Any]`` and a misspelled key only failed
at ``save()``. Each ``TypedDict`` here mirrors one ``KEYS`` array on the Rust
side, which is what actually rejects an unknown key;
``tests/test_option_types.py`` compares the two so they cannot drift apart.

Every one is ``total=False``: no key is required at the type level, because
which ones a rule needs depends on its ``type``, and Rust reports that with a
message naming the rule kind.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Dict, List, Tuple, TypedDict, Union

if TYPE_CHECKING:
    from .rustpy_xlsxwriter import Format

__all__ = [
    "Chart",
    "ChartSeries",
    "ConditionalRule",
    "DataValidation",
    "Image",
    "Margins",
    "Note",
    "Outline",
    "OutlineColumnGroup",
    "OutlineRowGroup",
    "PageSetup",
    "SheetView",
    "Sparkline",
]


class Margins(TypedDict, total=False):
    """Page margins in inches. Anything left out keeps Excel's default."""

    left: float
    right: float
    top: float
    bottom: float
    header: float
    footer: float


class PageSetup(TypedDict, total=False):
    """Page and print setup, passed as ``page_setup``."""

    landscape: bool
    paper_size: int
    margins: Margins
    print_area: Tuple[int, int, int, int]
    repeat_rows: Union[int, Tuple[int, int]]
    repeat_columns: Union[int, Tuple[int, int]]
    fit_to_pages: Tuple[int, int]
    scale: int
    center_horizontally: bool
    center_vertically: bool
    print_gridlines: bool
    print_headings: bool
    first_page_number: int
    header: str
    footer: str


class SheetView(TypedDict, total=False):
    """How Excel opens the sheet, passed as ``sheet_view``."""

    tab_color: str
    gridlines: bool
    zoom: int
    right_to_left: bool
    hidden: bool
    selected: bool


# ``from`` is a keyword, so these two need the functional form.
OutlineRowGroup = TypedDict(
    "OutlineRowGroup",
    {"from": int, "to": int, "collapsed": bool},
    total=False,
)
"""One row group: bounds are 0-based sheet row indices."""

OutlineColumnGroup = TypedDict(
    "OutlineColumnGroup",
    {"from": str, "to": str, "collapsed": bool},
    total=False,
)
"""One column group: bounds are header names."""


class Outline(TypedDict, total=False):
    """Collapsible row and column groups, passed as ``outline``."""

    rows: List[OutlineRowGroup]
    columns: List[OutlineColumnGroup]
    symbols_above: bool
    symbols_to_left: bool


class Note(TypedDict, total=False):
    """A cell note. A plain string is the short form for ``{"text": ...}``."""

    text: str
    author: str
    width: float
    height: float
    visible: bool
    background_color: str


class Image(TypedDict, total=False):
    """One anchored image. Give either ``path`` or ``data``, not both."""

    path: Union[str, os.PathLike]
    data: bytes
    row: int
    col: int
    scale: float
    scale_x: float
    scale_y: float
    fit_to_cell: bool
    keep_aspect_ratio: bool
    alt_text: str
    url: str


class ChartSeries(TypedDict, total=False):
    """One series. A plain column name is the short form for ``{"values": ...}``."""

    values: str
    name: str


class Chart(TypedDict, total=False):
    """One chart anchored to a cell."""

    type: str
    series: List[Union[str, ChartSeries]]
    categories: str
    row: int
    col: int
    title: str
    x_axis: str
    y_axis: str
    width: int
    height: int
    style: int
    legend: bool


Sparkline = TypedDict(
    "Sparkline",
    {
        "from": str,
        "to": str,
        "type": str,
        "color": str,
        "style": int,
        "high_point": bool,
        "low_point": bool,
        "first_point": bool,
        "last_point": bool,
        "markers": bool,
        "negative_points": bool,
        "axis": bool,
        "right_to_left": bool,
    },
    total=False,
)
"""One sparkline. ``from`` and ``to`` are the header names it spans."""


class DataValidation(TypedDict, total=False):
    """One validation rule. Which keys apply depends on ``type``."""

    type: str
    criteria: str
    value: float
    min: float
    max: float
    values: List[str]
    formula: str
    input_title: str
    input_message: str
    error_title: str
    error_message: str
    error_style: str
    show_dropdown: bool
    ignore_blank: bool


class ConditionalRule(TypedDict, total=False):
    """One conditional-format rule. Which keys apply depends on ``type``."""

    type: str
    criteria: str
    value: float
    min: float
    max: float
    format: "Format"
    color: str
    bar_only: bool
    min_color: str
    mid_color: str
    max_color: str


# ---------------------------------------------------------------------------
# Collections, as the options actually arrive
# ---------------------------------------------------------------------------

Notes = Dict[str, Union[str, Note]]
"""Notes keyed by column name."""

Images = List[Image]
"""Images anchored to a cell."""

Charts = List[Chart]
"""Charts anchored to a cell."""

Sparklines = Dict[str, Sparkline]
"""Sparklines keyed by the column they are drawn in."""

DataValidations = Dict[str, DataValidation]
"""Validation rules keyed by column name."""

ConditionalFormats = Dict[str, Union[ConditionalRule, List[ConditionalRule]]]
"""Conditional formats keyed by column name."""
