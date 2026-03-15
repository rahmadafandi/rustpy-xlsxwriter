"""
RustPy-XlsxWriter
==================

High-performance Excel file generation powered by Rust.

Quick start::

    from rustpy_xlsxwriter import FastExcel

    # Simple usage
    writer = FastExcel("output.xlsx")
    writer.sheet("Users", [{"Name": "Alice", "Age": 30}])
    writer.sheet("Items", [{"SKU": "A1", "Price": 9.99}])
    writer.save()

    # One-liner
    FastExcel("output.xlsx").sheet("Sheet1", records).save()

    # With options
    (
        FastExcel("report.xlsx", password="secret")
        .format(float_format="0.00", index_columns=["Name"])
        .freeze(row=1, col=1)
        .sheet("Data", records)
        .save()
    )

    # Pandas DataFrame
    FastExcel("df.xlsx").sheet("Sheet1", df).save()

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
    Union,
)

from .rustpy_xlsxwriter import (
    get_authors,
    get_description,
    get_homepage,
    get_license,
    get_name,
    get_repository,
    get_version,
    validate_sheet_name,
    write_worksheet,
    write_worksheets,
)

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
                Set to ``False`` for large datasets to improve performance.
        """
        self._target = target
        self._password = password
        self._autofit = autofit
        self._sheets: List[Dict[str, Any]] = []
        self._float_format: Optional[str] = None
        self._index_columns: Optional[List[str]] = None
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
        index_columns: Optional[List[str]] = None,
    ) -> "FastExcel":
        """Set number formatting and index column styling.

        Args:
            float_format: Excel number format for floats (e.g. ``"0.00"``).
            index_columns: Column names to render **bold**.
        """
        if float_format is not None:
            self._float_format = float_format
        if index_columns is not None:
            self._index_columns = index_columns
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
            ValueError: If the sheet name is invalid.
        """
        if not validate_sheet_name(name):
            raise ValueError(
                f"Invalid sheet name '{name}'. "
                "Must be ≤ 31 chars and cannot contain [ ] : * ? / \\"
            )
        self._sheets.append({name: data})
        return self

    # -- output -------------------------------------------------------------

    def save(self) -> None:
        """Write all sheets to the target file or buffer.

        Raises:
            ValueError: If no sheets have been added.
            OSError: If there are filesystem errors while writing.
        """
        if not self._sheets:
            raise ValueError("No sheets added. Call .sheet() before .save().")

        if len(self._sheets) == 1:
            sheet_name, data = next(iter(self._sheets[0].items()))
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
                index_columns=self._index_columns,
                autofit=self._autofit,
            )
        else:
            # Multi-sheet path
            write_worksheets(
                self._sheets,
                self._target,
                password=self._password,
                freeze_panes=self._freeze_panes or None,
                float_format=self._float_format,
                index_columns=self._index_columns,
                autofit=self._autofit,
            )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    # Class API
    "FastExcel",
    # Functional API
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
