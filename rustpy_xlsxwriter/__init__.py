"""
RustPy-XlsxWriter
==================

High-performance Excel file generation powered by Rust. ~9x faster than
Python's xlsxwriter.

Quick start::

    from rustpy_xlsxwriter import FastExcel

    # One-liner
    FastExcel("output.xlsx").sheet("Sheet1", records).save()

    # Multiple sheets with options
    (
        FastExcel("report.xlsx", password="secret")
        .format(float_format="0.00", index_columns=["Name"], bold_headers=True)
        .freeze(row=1, col=1)
        .sheet("Users", user_records)
        .sheet("Orders", order_records)
        .save()
    )

    # Context manager (auto-saves on exit)
    with FastExcel("output.xlsx") as f:
        f.sheet("Users", user_records)
        f.sheet("Orders", order_records)

    # Pandas DataFrame
    FastExcel("df.xlsx").sheet("Sheet1", pandas_df).save()

    # Polars DataFrame
    FastExcel("df.xlsx").sheet("Sheet1", polars_df).save()

    # In-memory buffer
    import io
    buf = io.BytesIO()
    FastExcel(buf).sheet("Sheet1", records).save()

    # Generator streaming (memory-efficient)
    def rows():
        for i in range(1_000_000):
            yield {"id": i, "value": f"row_{i}"}

    FastExcel("big.xlsx").sheet("Data", rows()).save()

You can also use the lower-level functional API directly::

    from rustpy_xlsxwriter import write_worksheet, write_worksheets
    write_worksheet([{"Name": "Alice"}], "output.xlsx")
"""

from __future__ import annotations

from typing import (
    Any,
    BinaryIO,
    Dict,
    List,
    Optional,
    Tuple,
    Union,
)

from importlib.metadata import metadata as _metadata
from importlib.metadata import version as _version

from .rustpy_xlsxwriter import (
    validate_sheet_name,
    write_csv,
    write_worksheet,
    write_worksheets,
)

_PKG = "rustpy-xlsxwriter"
_META = _metadata(_PKG)


def _project_url(label: str) -> str:
    prefix = f"{label}, "
    for entry in _META.get_all("Project-URL") or ():
        if entry.startswith(prefix):
            return entry[len(prefix):]
    return ""


def get_version() -> str:
    """Return the package version string."""
    return _version(_PKG)


def get_name() -> str:
    """Return the package name."""
    return _PKG


def get_authors() -> str:
    """Return the package authors (``'Name <email>'`` form)."""
    return _META.get("Author-email") or _META.get("Author") or ""


def get_description() -> str:
    """Return the package description."""
    return _META.get("Summary") or ""


def get_repository() -> str:
    """Return the repository URL."""
    return _project_url("Repository") or _META.get("Home-page") or ""


def get_homepage() -> str:
    """Return the homepage URL."""
    return _project_url("Homepage") or _META.get("Home-page") or ""


def get_license() -> str:
    """Return the license identifier."""
    return _META.get("License") or ""


__version__ = get_version()


# ---------------------------------------------------------------------------
# Builder-style class wrapper
# ---------------------------------------------------------------------------


class FastExcel:
    """Fluent builder for creating Excel files.

    Examples::

        # Minimal
        FastExcel("out.xlsx").sheet("Sheet1", records).save()

        # Full options
        (
            FastExcel("report.xlsx", password="s3cret")
            .format(float_format="0.00", index_columns=["ID"])
            .freeze(row=1)
            .sheet("Users", user_records)
            .sheet("Orders", order_records)
            .save()
        )
    """

    def __init__(
        self,
        target: Union[str, BinaryIO],
        *,
        password: Optional[str] = None,
        autofit: bool = True,
    ) -> None:
        """Create a new writer.

        Args:
            target: File path (``str``) or writable binary buffer
                (e.g. ``io.BytesIO``).
            password: Optional password to protect the workbook.
            autofit: Automatically adjust column widths (default ``True``).
                Under constant-memory mode (used by this library for all
                Excel paths) autofit sizing is approximate. Set to
                ``False`` for large datasets to improve performance.
        """
        self._target = target
        self._password = password
        self._autofit = autofit
        self._sheets: List[Tuple[str, Any]] = []
        self._float_format: Optional[str] = None
        self._datetime_format: Optional[str] = None
        self._index_columns: Optional[List[str]] = None
        self._bold_headers: bool = False
        self._freeze_panes: Dict[str, Dict[str, int]] = {}

    def __enter__(self) -> "FastExcel":
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        if exc_type is None and self._sheets:
            self.save()

    # -- configuration (chainable) ------------------------------------------

    def format(
        self,
        *,
        float_format: Optional[str] = None,
        datetime_format: Optional[str] = None,
        index_columns: Optional[List[str]] = None,
        bold_headers: Optional[bool] = None,
    ) -> "FastExcel":
        """Set number formatting and column styling.

        Args:
            float_format: Excel number format for floats (e.g. ``"0.00"``).
            datetime_format: Excel number format for datetimes
                (default ``"yyyy-mm-ddThh:mm:ss"``).
            index_columns: Column names to render **bold**.
            bold_headers: Whether to render header row in **bold**.
        """
        if float_format is not None:
            self._float_format = float_format
        if datetime_format is not None:
            self._datetime_format = datetime_format
        if index_columns is not None:
            self._index_columns = index_columns
        if bold_headers is not None:
            self._bold_headers = bold_headers
        return self

    def freeze(
        self,
        *,
        row: Optional[int] = None,
        col: Optional[int] = None,
        sheet: Optional[str] = None,
    ) -> "FastExcel":
        """Configure freeze panes.

        Args:
            row: Freeze panes above this row number.
            col: Freeze panes to the left of this column number.
            sheet: Apply to a specific sheet only. If ``None``, applies
                to all sheets (``"general"``).
        """
        key = sheet or "general"
        config: Dict[str, int] = {}
        if row is not None:
            config["row"] = row
        if col is not None:
            config["col"] = col
        if config:
            self._freeze_panes[key] = config
        return self

    # -- data ---------------------------------------------------------------

    def sheet(self, name: str, data: Any) -> "FastExcel":
        """Add a worksheet with data.

        Args:
            name: Sheet name (≤ 31 chars, no ``[ ] : * ? / \\``).
            data: List of dicts, generator of dicts, or pandas DataFrame.

        Raises:
            ValueError: If the sheet name is invalid (validated on save).
        """
        self._sheets.append((name, data))
        return self

    # -- output -------------------------------------------------------------

    def save(self) -> None:
        """Write all sheets to the target file or buffer.

        Automatically detects output format from file extension:
        - ``.xlsx`` → Excel (default)
        - ``.csv`` → CSV
        - ``.tsv`` → TSV (tab-separated)

        Raises:
            ValueError: If no sheets have been added.
            OSError: If there are filesystem errors while writing.
        """
        if not self._sheets:
            raise ValueError("No sheets added. Call .sheet() before .save().")

        # Auto-detect CSV/TSV from file extension
        if isinstance(self._target, str):
            lower = self._target.lower()
            if lower.endswith(".csv") or lower.endswith(".tsv"):
                if len(self._sheets) > 1:
                    raise ValueError(
                        f"CSV/TSV output supports a single sheet; got {len(self._sheets)}."
                    )
                delimiter = "\t" if lower.endswith(".tsv") else ","
                _, data = self._sheets[0]
                write_csv(data, self._target, delimiter=delimiter)
                return

        if len(self._sheets) == 1:
            sheet_name, data = self._sheets[0]
            # Single-sheet path: use write_worksheet for simpler freeze pane
            freeze_row = None
            freeze_col = None
            # Check general or sheet-specific freeze config
            cfg = self._freeze_panes.get(sheet_name) or self._freeze_panes.get(
                "general"
            )
            if cfg:
                freeze_row = cfg.get("row")
                freeze_col = cfg.get("col")

            write_worksheet(
                data,
                self._target,
                sheet_name=sheet_name,
                password=self._password,
                freeze_row=freeze_row,
                freeze_col=freeze_col,
                float_format=self._float_format,
                datetime_format=self._datetime_format,
                index_columns=self._index_columns,
                autofit=self._autofit,
                bold_headers=self._bold_headers,
            )
        else:
            # Multi-sheet path
            write_worksheets(
                self._sheets,
                self._target,
                password=self._password,
                freeze_panes=self._freeze_panes or None,
                float_format=self._float_format,
                datetime_format=self._datetime_format,
                index_columns=self._index_columns,
                autofit=self._autofit,
                bold_headers=self._bold_headers,
            )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    # Class API
    "FastExcel",
    # Functional API
    "write_csv",
    "write_worksheet",
    "write_worksheets",
    # Utilities
    "validate_sheet_name",
    # Metadata
    "get_version",
    "get_name",
    "get_authors",
    "get_description",
    "get_repository",
    "get_homepage",
    "get_license",
    # Convenience
    "__version__",
]
