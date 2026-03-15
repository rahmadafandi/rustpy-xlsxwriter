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

FileTarget = Union[str, BinaryIO]
"""A file path (``str``) or a writable binary buffer (e.g. ``io.BytesIO``)."""

FreezePanesConfig = Dict[str, Dict[str, int]]
"""Freeze-pane configuration.

Example::

    {
        "general":  {"row": 1, "col": 0},   # applies to every sheet
        "Sheet1":   {"row": 1, "col": 2},   # override for Sheet1
    }
"""

SheetData = Union[Records, DataFrame]
"""Data accepted per sheet – either :data:`Records` or a :data:`DataFrame`."""

SheetMap = Dict[str, SheetData]
"""Maps a sheet name to its data, e.g. ``{"Sheet1": records}``."""

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
    ) -> None:
        """Create a new writer.

        Args:
            target: File path or writable binary buffer (e.g. ``io.BytesIO``).
            password: Optional password to protect the workbook.
            autofit: Automatically adjust column widths (default ``True``).
                Set to ``False`` for large datasets to improve performance.
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

    def sheet(self, name: str, data: SheetData) -> FastExcel:
        """Add a worksheet with data.

        Args:
            name: Sheet name (≤ 31 chars, no ``[ ] : * ? / \\``).
            data: List of dicts, generator of dicts, pandas DataFrame,
                or polars DataFrame.

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

    Raises:
        ValueError: Invalid sheet name or unsupported data type.
        OSError: File system error while writing.

    Examples:
        >>> write_worksheet([{"Name": "Alice", "Age": 30}], "out.xlsx")
    """
    ...

def write_worksheets(
    records_with_sheet_name: List[SheetMap],
    file_name: FileTarget,
    password: Optional[str] = None,
    freeze_panes: Optional[FreezePanesConfig] = None,
    float_format: Optional[str] = None,
    datetime_format: Optional[str] = None,
    index_columns: Optional[List[str]] = None,
    autofit: bool = True,
    bold_headers: bool = False,
) -> None:
    """Write data to **multiple** worksheets in an Excel file.

    Args:
        records_with_sheet_name: A list where each element is a dict
            mapping **one** sheet name to its data.
        file_name: Destination file path or writable binary buffer.
        password: Optional password to protect the workbook.
        freeze_panes: Per-sheet and/or general freeze-pane config.
        float_format: Excel number format for floats (e.g. ``"0.00"``).
        index_columns: Column names that should be rendered **bold**.
        autofit: Automatically adjust column widths (default ``True``).

    Raises:
        ValueError: Invalid sheet name or unsupported data type.
        OSError: File system error while writing.

    Examples:
        >>> write_worksheets(
        ...     [{"Users": [{"Name": "Alice"}]}, {"Items": [{"SKU": "A1"}]}],
        ...     "multi.xlsx",
        ... )
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
    """Return the package version string (e.g. ``'0.1.0'``)."""
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

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    "FastExcel",
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
    "SheetData",
    "SheetMap",
]
