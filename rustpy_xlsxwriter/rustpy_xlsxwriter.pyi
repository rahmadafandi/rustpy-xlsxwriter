"""Type stubs for the compiled extension module.

Covers only what ``lib.rs`` exports. ``FastExcel`` and the metadata
helpers live in ``__init__.py`` and are annotated inline there.
"""

from __future__ import annotations

from typing import (
    Any,
    BinaryIO,
    Dict,
    Iterable,
    List,
    Optional,
    Tuple,
    Union,
)

# ---------------------------------------------------------------------------
# Type aliases
# ---------------------------------------------------------------------------

Record = Dict[str, Any]
"""A single row of data represented as ``{column_name: value}``."""

Records = Union[List[Record], Iterable[Record]]
"""A list (or any iterable, including generators) of :data:`Record` dicts."""

DataFrame = Any
"""A *pandas* or *polars* ``DataFrame`` – kept as ``Any`` to avoid a hard dependency."""

import os as _os

FileTarget = Union[str, _os.PathLike, BinaryIO]
"""A file path (``str`` / :class:`os.PathLike`) or a writable binary buffer (e.g. ``io.BytesIO``)."""

FreezePanesConfig = Dict[str, Dict[str, int]]
"""Freeze-pane configuration.

Example::

    {
        "general":  {"row": 1, "col": 0},   # applies to every sheet
        "Sheet1":   {"row": 1, "col": 2},   # override for Sheet1
    }
"""

ColumnWidths = Union[Dict[str, float], List[float]]
"""Per-column width — a dict keyed by header name or a positional list of widths."""

ColumnFormats = Union[Dict[str, "Format"], List["Format"]]
"""Per-column formats — a dict keyed by header name or a positional list of :class:`Format`."""

MergeRange = Union[
    Tuple[int, int, int, int, Any],
    Tuple[int, int, int, int, Any, Optional["Format"]],
]
"""One merged cell range: ``(first_row, first_col, last_row, last_col, value)``,
optionally followed by a :class:`Format`."""

SheetData = Union[Records, DataFrame]
"""Data accepted per sheet – either :data:`Records` or a :data:`DataFrame`."""

SheetMap = Dict[str, SheetData]
"""(Legacy alias) Maps a sheet name to its data."""

SheetEntry = Tuple[str, SheetData]
"""A ``(sheet_name, data)`` pair as accepted by :func:`write_worksheets`."""

# ---------------------------------------------------------------------------
# Cell format
# ---------------------------------------------------------------------------

class Format:
    """A reusable cell format (font, fill, border, alignment, number format).

    Setters are chainable — each returns ``self``::

        Format().set_bold().set_font_color("#FF0000").set_num_format("0.00%")

    Colors accept ``"#RRGGBB"`` / ``"RRGGBB"`` hex or a color name
    (e.g. ``"red"``). Enum-valued setters accept lowercase string names.
    """

    def __init__(self) -> None: ...
    # Font
    def set_bold(self) -> Format: ...
    def set_italic(self) -> Format: ...
    def set_underline(self, style: str = "single", /) -> Format: ...
    def set_font_strikethrough(self) -> Format: ...
    def set_font_size(self, size: float, /) -> Format: ...
    def set_font_name(self, name: str, /) -> Format: ...
    def set_font_color(self, color: str, /) -> Format: ...
    def set_font_script(self, script: str, /) -> Format: ...
    def set_font_family(self, n: int, /) -> Format: ...
    def set_font_charset(self, n: int, /) -> Format: ...
    def set_font_scheme(self, scheme: str, /) -> Format: ...
    # Fill
    def set_background_color(self, color: str, /) -> Format: ...
    def set_foreground_color(self, color: str, /) -> Format: ...
    def set_pattern(self, pattern: str, /) -> Format: ...
    # Border
    def set_border(self, style: str, /) -> Format: ...
    def set_border_color(self, color: str, /) -> Format: ...
    def set_border_top(self, style: str, /) -> Format: ...
    def set_border_bottom(self, style: str, /) -> Format: ...
    def set_border_left(self, style: str, /) -> Format: ...
    def set_border_right(self, style: str, /) -> Format: ...
    def set_border_top_color(self, color: str, /) -> Format: ...
    def set_border_bottom_color(self, color: str, /) -> Format: ...
    def set_border_left_color(self, color: str, /) -> Format: ...
    def set_border_right_color(self, color: str, /) -> Format: ...
    def set_border_diagonal(self, style: str, /) -> Format: ...
    def set_border_diagonal_color(self, color: str, /) -> Format: ...
    def set_border_diagonal_type(self, t: str, /) -> Format: ...
    # Alignment / layout
    def set_align(self, align: str, /) -> Format: ...
    def set_text_wrap(self) -> Format: ...
    def set_rotation(self, degrees: int, /) -> Format: ...
    def set_indent(self, n: int, /) -> Format: ...
    def set_shrink(self) -> Format: ...
    def set_reading_direction(self, n: int, /) -> Format: ...
    # Number
    def set_num_format(self, fmt: str, /) -> Format: ...
    def set_num_format_index(self, i: int, /) -> Format: ...
    # Protection / misc
    def set_locked(self) -> Format: ...
    def set_unlocked(self) -> Format: ...
    def set_hidden(self) -> Format: ...
    def set_quote_prefix(self) -> Format: ...
    def set_checkbox(self) -> Format: ...
    def set_hyperlink(self) -> Format: ...

# ---------------------------------------------------------------------------
# Core write functions
# ---------------------------------------------------------------------------

def write_worksheet(
    records: SheetData,
    file_name: FileTarget,
    sheet_name: Optional[str] = None,
    password: Optional[str] = None,
    freeze_row: Optional[int] = None,
    freeze_col: Optional[int] = None,
    float_format: Optional[str] = None,
    datetime_format: Optional[str] = None,
    index_columns: Optional[List[str]] = None,
    autofit: bool = True,
    bold_headers: bool = False,
    column_width: Optional[float] = None,
    column_widths: Optional[ColumnWidths] = None,
    column_formats: Optional[ColumnFormats] = None,
    header_format: Optional[Format] = None,
    dedupe_strings: bool = False,
    header_row: int = 0,
    merge_ranges: Optional[List[MergeRange]] = None,
    row_heights: Optional[Dict[int, float]] = None,
    row_formats: Optional[Dict[int, Format]] = None,
    banded_rows: Optional[str] = None,
    autofilter: bool = False,
    url_columns: Optional[List[str]] = None,
    totals_row: Optional[Dict[str, str]] = None,
    totals_label: Optional[str] = None,
    totals_format: Optional[Format] = None,
    formula_columns: Optional[Dict[str, str]] = None,
) -> None:
    """Write data to a **single** worksheet in an Excel file.

    Args:
        records: Data to write – a list of dicts, a generator of dicts,
            or a *pandas* ``DataFrame``.
        file_name: Destination file path (``*.xlsx``) **or** a writable
            binary buffer such as ``io.BytesIO``.
        sheet_name: Worksheet name (default ``"Sheet1"``).
            Must be ≤ 31 chars; cannot contain ``[ ] : * ? / \\``.
        password: Optional password to protect the workbook.
        freeze_row: Freeze panes above this row number.
        freeze_col: Freeze panes to the left of this column number.
        float_format: Excel number format for floats (e.g. ``"0.00"``).
        index_columns: Column names that should be rendered **bold**.
        autofit: Automatically adjust column widths (default ``True``).
        column_width: Uniform width applied to every column.
        column_widths: Per-column width — a dict keyed by header name or a positional list.
        dedupe_strings: Store repeated strings once in the shared-string table
            instead of inline. Shrinks files with heavily repeated text, at the
            cost of buffering the sheet in memory (disables constant-memory
            mode). Off by default.
        header_row: 0-based row the header is written on; data follows it.
        merge_ranges: ``(first_row, first_col, last_row, last_col, value[, format])``
            tuples. Must sit strictly above ``header_row``.
        row_heights: ``{row_index: height}`` in points.
        row_formats: ``{row_index: Format}`` applied to the whole row.
        banded_rows: Background colour shaded onto every other data row.
        autofilter: Add filter dropdowns over the header row and its data.
        url_columns: Column names whose text cells become clickable links.
            Values Excel rejects fall back to plain text.
        totals_row: ``{column_name: aggregate}`` written as formulas below the
            data. Valid: sum, average, count, min, max, product, stdev.
        totals_label: Text for the first column of the totals row.
        totals_format: Format applied to the whole totals row.
        formula_columns: ``{header: formula}`` appended after the data, one
            formula per row. ``{row}``/``{first}`` are substituted.

    Raises:
        ValueError: Invalid sheet name or unsupported data type.
        OSError: File system error while writing.

    Examples:
        >>> write_worksheet([{"Name": "Alice", "Age": 30}], "out.xlsx")
    """
    ...

def write_worksheets(
    records_with_sheet_name: List[SheetEntry],
    file_name: FileTarget,
    password: Optional[str] = None,
    freeze_panes: Optional[FreezePanesConfig] = None,
    float_format: Optional[str] = None,
    datetime_format: Optional[str] = None,
    index_columns: Optional[List[str]] = None,
    autofit: bool = True,
    bold_headers: bool = False,
    column_width: Optional[Dict[str, float]] = None,
    column_widths: Optional[Dict[str, ColumnWidths]] = None,
    column_formats: Optional[Dict[str, ColumnFormats]] = None,
    header_format: Optional[Dict[str, Format]] = None,
    dedupe_strings: Optional[Dict[str, bool]] = None,
    header_row: Optional[Dict[str, int]] = None,
    merge_ranges: Optional[Dict[str, List[MergeRange]]] = None,
    row_heights: Optional[Dict[str, Dict[int, float]]] = None,
    row_formats: Optional[Dict[str, Dict[int, Format]]] = None,
    banded_rows: Optional[Dict[str, str]] = None,
    autofilter: Optional[Dict[str, bool]] = None,
    url_columns: Optional[Dict[str, List[str]]] = None,
    totals_row: Optional[Dict[str, Dict[str, str]]] = None,
    totals_label: Optional[Dict[str, str]] = None,
    totals_format: Optional[Dict[str, Format]] = None,
    formula_columns: Optional[Dict[str, Dict[str, str]]] = None,
) -> None:
    """Write data to **multiple** worksheets in an Excel file.

    Args:
        records_with_sheet_name: A list of ``(sheet_name, data)`` tuples.
        file_name: Destination file path or writable binary buffer.
        password: Optional password to protect the workbook.
        freeze_panes: Per-sheet and/or general freeze-pane config.
        float_format: Excel number format for floats (e.g. ``"0.00"``).
        index_columns: Column names that should be rendered **bold**.
        autofit: Automatically adjust column widths (default ``True``).
        column_width: Uniform width per sheet — dict keyed by sheet name (``"general"`` applies to all).
        column_widths: Per-column width per sheet — dict keyed by sheet name mapping to :data:`ColumnWidths`.
        dedupe_strings: Per-sheet shared-string deduplication — dict keyed by
            sheet name (``"general"`` applies to all). See
            :func:`write_worksheet` for the trade-off.
        header_row: Per-sheet header row index — dict keyed by sheet name.
        merge_ranges: Per-sheet merged cells — dict keyed by sheet name.
        row_heights: Per-sheet row heights — dict keyed by sheet name.
        row_formats: Per-sheet row formats — dict keyed by sheet name.
        banded_rows: Per-sheet alternating row colour — dict keyed by sheet name.
        autofilter: Per-sheet filter dropdowns — dict keyed by sheet name.
        url_columns: Per-sheet link columns — dict keyed by sheet name.
        totals_row: Per-sheet totals formulas — dict keyed by sheet name.
        totals_label: Per-sheet totals label — dict keyed by sheet name.
        totals_format: Per-sheet totals row format — dict keyed by sheet name.
        formula_columns: Per-sheet computed columns — dict keyed by sheet name.

    Raises:
        ValueError: Invalid sheet name or unsupported data type.
        OSError: File system error while writing.

    Examples:
        >>> write_worksheets(
        ...     [("Users", [{"Name": "Alice"}]), ("Items", [{"SKU": "A1"}])],
        ...     "multi.xlsx",
        ... )
    """
    ...

# ---------------------------------------------------------------------------
# CSV writer
# ---------------------------------------------------------------------------

def write_csv(
    records: SheetData,
    file_name: FileTarget,
    delimiter: Optional[str] = None,
    sanitize_formulas: bool = False,
    bom: bool = False,
    columns: Optional[List[str]] = None,
    header: bool = True,
) -> None:
    """Write data to a CSV file.

    Args:
        records: Data to write – a list of dicts, a generator of dicts,
            a *pandas* ``DataFrame``, or a *polars* ``DataFrame``.
        file_name: Destination file path or writable binary buffer.
        delimiter: Column delimiter (default ``","``). Use ``"\\t"`` for TSV.
        sanitize_formulas: When ``True``, string fields starting with
            ``= + - @`` are prefixed with ``'`` to neutralize CSV formula
            injection. Off by default (output stays byte-identical).
        bom: Prefix the UTF-8 byte order mark. Excel on Windows reads a BOM-less
            UTF-8 file as the system code page, which turns non-ASCII text into
            mojibake; this is the fix. Off by default so output stays
            byte-identical for pipelines that parse it.
        columns: Select and order the output columns by name. A name that is
            not in the data raises ``ValueError`` — unlike the styling options,
            this one decides the shape of the file, so a silent drop would hand
            back something that looks complete. Works on every input type,
            and stays zero-copy on the Arrow path.
        header: Write the header row. Set ``False`` to append to an existing
            file or to feed a reader that supplies its own names.

    Examples:
        >>> write_csv([{"Name": "Alice", "Age": 30}], "out.csv")
    """
    ...

# ---------------------------------------------------------------------------
# Sheet-name validation
# ---------------------------------------------------------------------------

def validate_sheet_name(name: str) -> bool:
    """Check whether *name* is a valid Excel sheet name.

    Rules: ≤ 31 characters, no ``[ ] : * ? / \\``, not empty.

    Examples:
        >>> validate_sheet_name("Sheet1")
        True
        >>> validate_sheet_name("Sheet[1]")
        False
    """
    ...
