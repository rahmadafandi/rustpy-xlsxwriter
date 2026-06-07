"""Type stubs for rustpy_xlsxwriter – high-performance Excel writer powered by Rust."""

from __future__ import annotations

from types import TracebackType
from typing import (
    Any,
    BinaryIO,
    Dict,
    Iterable,
    List,
    Optional,
    Tuple,
    Type,
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
    def set_underline(self, style: str = "single") -> Format: ...
    def set_font_strikethrough(self) -> Format: ...
    def set_font_size(self, size: float) -> Format: ...
    def set_font_name(self, name: str) -> Format: ...
    def set_font_color(self, color: str) -> Format: ...
    def set_font_script(self, script: str) -> Format: ...
    def set_font_family(self, n: int) -> Format: ...
    def set_font_charset(self, n: int) -> Format: ...
    def set_font_scheme(self, scheme: str) -> Format: ...
    # Fill
    def set_background_color(self, color: str) -> Format: ...
    def set_foreground_color(self, color: str) -> Format: ...
    def set_pattern(self, pattern: str) -> Format: ...
    # Border
    def set_border(self, style: str) -> Format: ...
    def set_border_color(self, color: str) -> Format: ...
    def set_border_top(self, style: str) -> Format: ...
    def set_border_bottom(self, style: str) -> Format: ...
    def set_border_left(self, style: str) -> Format: ...
    def set_border_right(self, style: str) -> Format: ...
    def set_border_top_color(self, color: str) -> Format: ...
    def set_border_bottom_color(self, color: str) -> Format: ...
    def set_border_left_color(self, color: str) -> Format: ...
    def set_border_right_color(self, color: str) -> Format: ...
    def set_border_diagonal(self, style: str) -> Format: ...
    def set_border_diagonal_color(self, color: str) -> Format: ...
    def set_border_diagonal_type(self, t: str) -> Format: ...
    # Alignment / layout
    def set_align(self, align: str) -> Format: ...
    def set_text_wrap(self) -> Format: ...
    def set_rotation(self, degrees: int) -> Format: ...
    def set_indent(self, n: int) -> Format: ...
    def set_shrink(self) -> Format: ...
    def set_reading_direction(self, n: int) -> Format: ...
    # Number
    def set_num_format(self, fmt: str) -> Format: ...
    def set_num_format_index(self, i: int) -> Format: ...
    # Protection / misc
    def set_locked(self) -> Format: ...
    def set_unlocked(self) -> Format: ...
    def set_hidden(self) -> Format: ...
    def set_quote_prefix(self) -> Format: ...
    def set_checkbox(self) -> Format: ...
    def set_hyperlink(self) -> Format: ...

# ---------------------------------------------------------------------------
# Builder class
# ---------------------------------------------------------------------------

class FastExcel:
    """Fluent builder for creating Excel files.

    Examples::

        FastExcel("out.xlsx").sheet("Sheet1", records).save()

        (
            FastExcel("report.xlsx", password="s3cret")
            .format(float_format="0.00", index_columns=["ID"])
            .freeze(row=1)
            .sheet("Users", user_records)
            .sheet("Orders", order_records)
            .save()
        )

        # Context manager (auto-saves on exit)
        with FastExcel("out.xlsx") as f:
            f.sheet("Sheet1", records)
    """

    def __init__(
        self,
        target: FileTarget,
        *,
        password: Optional[str] = None,
        autofit: bool = True,
        sanitize_formulas: bool = False,
    ) -> None:
        """Create a new writer.

        Args:
            target: File path or writable binary buffer (e.g. ``io.BytesIO``).
            password: Optional worksheet-protection password. Sets Excel's sheet
                protection flag only — it does NOT encrypt the file; data is
                stored in plaintext.
            autofit: Automatically adjust column widths (default ``True``).
                Set to ``False`` for large datasets to improve performance.
            sanitize_formulas: CSV/TSV only. When ``True``, string fields
                starting with ``= + - @`` are prefixed with ``'`` to neutralize
                CSV formula injection. Off by default. No effect on ``.xlsx``.
        """
        ...

    def __enter__(self) -> FastExcel: ...

    def __exit__(
        self,
        exc_type: Optional[Type[BaseException]],
        exc_val: Optional[BaseException],
        exc_tb: Optional[TracebackType],
    ) -> None: ...

    def format(
        self,
        *,
        float_format: Optional[str] = None,
        datetime_format: Optional[str] = None,
        index_columns: Optional[List[str]] = None,
        bold_headers: Optional[bool] = None,
    ) -> FastExcel:
        """Set number formatting and column styling.

        Args:
            float_format: Excel number format for floats (e.g. ``"0.00"``).
            datetime_format: Excel number format for datetimes
                (default ``"yyyy-mm-ddThh:mm:ss"``).
            index_columns: Column names to render **bold**.
            bold_headers: Whether to render header row in **bold**.
        """
        ...

    def freeze(
        self,
        *,
        row: Optional[int] = None,
        col: Optional[int] = None,
        sheet: Optional[str] = None,
    ) -> FastExcel:
        """Configure freeze panes.

        Args:
            row: Freeze panes above this row number.
            col: Freeze panes to the left of this column number.
            sheet: Apply to a specific sheet. If ``None``, applies to all.
        """
        ...

    def sheet(
        self,
        name: str,
        data: SheetData,
        *,
        column_width: Optional[float] = None,
        column_widths: Optional[ColumnWidths] = None,
        column_formats: Optional[ColumnFormats] = None,
        header_format: Optional[Format] = None,
    ) -> FastExcel:
        """Add a worksheet with data.

        Args:
            name: Sheet name (≤ 31 chars, no ``[ ] : * ? / \\``).
            data: List of dicts, generator of dicts, pandas DataFrame,
                or polars DataFrame.
            column_width: Uniform width applied to every column of this sheet.
            column_widths: Per-column width — a dict keyed by header name
                or a positional list of widths.
            column_formats: Per-column :class:`Format` — a dict keyed by header
                name or a positional list.
            header_format: A :class:`Format` applied to the header row.

        Raises:
            ValueError: If the sheet name is invalid.
        """
        ...

    def save(self) -> None:
        """Write all sheets to the target file or buffer.

        Raises:
            ValueError: If no sheets have been added.
            OSError: File system error while writing.
        """
        ...

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

# ---------------------------------------------------------------------------
# Package metadata
# ---------------------------------------------------------------------------

def get_version() -> str:
    """Return the package version string (e.g. ``'0.5.2'``)."""
    ...

def get_name() -> str:
    """Return the package name (``'rustpy-xlsxwriter'``)."""
    ...

def get_authors() -> str:
    """Return the package authors."""
    ...

def get_description() -> str:
    """Return the package description."""
    ...

def get_repository() -> str:
    """Return the repository URL."""
    ...

def get_homepage() -> str:
    """Return the homepage URL."""
    ...

def get_license() -> str:
    """Return the license identifier (``'MIT'``)."""
    ...

__version__: str
"""Package version string — same value as :func:`get_version`."""

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    "FastExcel",
    "Format",
    "write_csv",
    "write_worksheet",
    "write_worksheets",
    "validate_sheet_name",
    "get_version",
    "get_name",
    "get_authors",
    "get_description",
    "get_repository",
    "get_homepage",
    "get_license",
    "Record",
    "Records",
    "DataFrame",
    "FileTarget",
    "FreezePanesConfig",
    "ColumnWidths",
    "ColumnFormats",
    "SheetData",
    "SheetEntry",
    "SheetMap",
    "__version__",
]
