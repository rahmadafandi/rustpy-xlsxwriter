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

import os as _os
import warnings as _warnings
from importlib.metadata import metadata as _metadata
from importlib.metadata import version as _version

from .rustpy_xlsxwriter import (
    Format,
    validate_sheet_name,
)
from .rustpy_xlsxwriter import write_csv as _write_csv_rs
from .rustpy_xlsxwriter import write_worksheet as _write_worksheet_rs
from .rustpy_xlsxwriter import write_worksheets as _write_worksheets_rs


def _coerce_target(target: Any) -> Any:
    """Accept str, bytes, or any os.PathLike as a file path; pass other
    objects (file-like buffers) through unchanged."""
    if isinstance(target, _os.PathLike):
        return _os.fspath(target)
    return target


def write_worksheet(records, file_name, *args, **kwargs):
    return _write_worksheet_rs(records, _coerce_target(file_name), *args, **kwargs)


def write_worksheets(records_with_sheet_name, file_name, *args, **kwargs):
    return _write_worksheets_rs(
        records_with_sheet_name, _coerce_target(file_name), *args, **kwargs
    )


def write_csv(records, file_name, *args, **kwargs):
    return _write_csv_rs(records, _coerce_target(file_name), *args, **kwargs)

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

#: Options ``sheet()`` records per sheet and forwards to the writers. Both save
#: paths iterate this, so adding an option means touching only ``sheet()``.
_PER_SHEET_OPTIONS = (
    "column_width",
    "column_widths",
    "column_formats",
    "header_format",
    "dedupe_strings",
    "header_row",
    "merge_ranges",
    "row_heights",
    "row_formats",
    "banded_rows",
    "autofilter",
    "url_columns",
    "totals_row",
    "totals_label",
    "totals_format",
    "formula_columns",
)


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
        target: Union[str, _os.PathLike, BinaryIO],
        *,
        password: Optional[str] = None,
        autofit: bool = True,
        sanitize_formulas: bool = False,
    ) -> None:
        """Create a new writer.

        Args:
            target: File path (``str`` or :class:`os.PathLike`, e.g.
                ``pathlib.Path``) or writable binary buffer
                (e.g. ``io.BytesIO``).
            password: Optional worksheet-protection password. NOTE: this sets
                Excel's *sheet protection* flag only — it does **not** encrypt
                the file. The cell data is stored in plaintext and the
                protection is trivially removed; do not rely on it to keep
                data confidential.
            autofit: Automatically adjust column widths (default ``True``).
                Under constant-memory mode (the default for every Excel sheet,
                unless ``sheet(..., dedupe_strings=True)`` opts out) autofit
                sizing is approximate. Set to ``False`` for large datasets to
                improve performance.
            sanitize_formulas: CSV/TSV only. When ``True``, string fields that
                begin with ``= + - @`` are prefixed with a single quote so
                spreadsheet apps open them as text instead of executing them as
                formulas (CSV-injection mitigation). Off by default to keep
                output byte-identical. Has no effect on ``.xlsx`` output, where
                values are already written as text cells.
        """
        self._target = _coerce_target(target)
        self._password = password
        self._autofit = autofit
        self._sanitize_formulas = sanitize_formulas
        self._sheets: List[Tuple[str, Any]] = []
        self._float_format: Optional[str] = None
        self._datetime_format: Optional[str] = None
        self._index_columns: Optional[List[str]] = None
        self._bold_headers: bool = False
        self._freeze_panes: Dict[str, Dict[str, int]] = {}
        # {option: {sheet_name: value}} — see _PER_SHEET_OPTIONS.
        self._per_sheet: Dict[str, Dict[str, Any]] = {
            option: {} for option in _PER_SHEET_OPTIONS
        }

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

    def sheet(
        self,
        name: str,
        data: Any,
        *,
        column_width: Optional[float] = None,
        column_widths: Optional[Union[Dict[str, float], List[float]]] = None,
        column_formats: Optional[Union[Dict[str, "Format"], List["Format"]]] = None,
        header_format: Optional["Format"] = None,
        dedupe_strings: bool = False,
        header_row: int = 0,
        merge_ranges: Optional[List[Tuple]] = None,
        row_heights: Optional[Dict[int, float]] = None,
        row_formats: Optional[Dict[int, "Format"]] = None,
        banded_rows: Optional[str] = None,
        autofilter: bool = False,
        url_columns: Optional[List[str]] = None,
        totals_row: Optional[Dict[str, str]] = None,
        totals_label: Optional[str] = None,
        totals_format: Optional["Format"] = None,
        formula_columns: Optional[Dict[str, str]] = None,
    ) -> "FastExcel":
        """Add a worksheet with data.

        Args:
            name: Sheet name (≤ 31 chars, no ``[ ] : * ? / \\``).
            data: List of dicts, generator of dicts, or pandas DataFrame.
            column_width: Uniform width applied to every column of this sheet.
            column_widths: Per-column width — a dict keyed by header name
                (``{"name": 22}``) or a positional list (``[7, 22, 40]``).
                Overrides ``column_width`` for the columns it names.
            column_formats: Per-column :class:`Format` — a dict keyed by header
                name (``{"name": Format().set_bold()}``) or a positional list
                (``[Format().set_bold(), None]``).
            header_format: :class:`Format` applied to every header cell of this
                sheet.
            dedupe_strings: Store repeated strings once in the workbook's
                shared-string table instead of inline, which can shrink the
                ``.xlsx`` substantially when a sheet has many repeated text
                values (categories, statuses, country codes). Off by default:
                it takes this sheet out of constant-memory mode, so the whole
                sheet is buffered in RAM and every string is hashed. Turn it on
                per sheet, for sheets whose text actually repeats, and measure.
            header_row: 0-based row the header is written on; data follows it.
                Raise it to leave room for merged banner headers above.
            merge_ranges: Merged cells, as
                ``(first_row, first_col, last_row, last_col, value[, format])``
                tuples — e.g. ``[(0, 1, 0, 2, "Gender", banner_fmt)]`` for a
                banner spanning two sub-columns. Ranges must sit strictly above
                ``header_row``; anything overlapping the header or data raises,
                because rows already written cannot be merged after the fact.
            row_heights: ``{row_index: height}`` in points.
            row_formats: ``{row_index: Format}`` applied to the whole row — the
                way to put a bottom border under the header or a top border
                above a totals row. A cell carrying its own format (a number
                format, a column format, a band) wins over the row's.
            banded_rows: Background colour (``"#F2F2F2"`` or a name) shaded onto
                every other data row, starting with the second. Applied per cell
                rather than per row, so columns with their own number format
                stay shaded too.
            autofilter: Add Excel's filter dropdowns over the header row and its
                data. The range is computed from the rows actually written, so
                it follows ``header_row`` and needs no manual bounds.
            url_columns: Column names whose text cells become clickable links —
                ``["homepage"]``. Accepts what Excel accepts: ``http(s)://``,
                ``mailto:``, and ``internal:Sheet2!A1`` for a link to another
                sheet. A value Excel would reject (ordinary text, or a URL past
                its 2083-character limit) is written as plain text instead, so a
                stray non-link never aborts the export. The cell displays the
                URL itself; per-cell display text is not supported.
            totals_row: ``{column_name: aggregate}`` written as Excel formulas
                in a row below the data — ``{"amount": "sum"}`` becomes
                ``=SUM(C2:C101)``. Valid aggregates: ``sum``, ``average``,
                ``count``, ``min``, ``max``, ``product``, ``stdev``. A value
                starting with ``=`` is used as a formula instead, with ``{col}``
                the column letter and ``{first}``/``{last}`` the data range::

                    totals_row={"margin": "=SUM({col}{first}:{col}{last})/2"}

                Skipped entirely when there are no data rows, since the range
                would be empty. NOTE: the formulas carry no computed result, so
                readers that use cached values (``pandas.read_excel``,
                ``openpyxl`` with ``data_only=True``) get ``None`` until Excel
                or LibreOffice opens the file and recalculates.
            totals_label: Text for the first column of the totals row, e.g.
                ``"Total"``. Raises if the first column also has an aggregate.
            totals_format: :class:`Format` for the whole totals row — the usual
                bold plus a top border. Needed because the row index is not
                known ahead of time, so ``row_formats`` cannot reach it.
            formula_columns: ``{header: formula}`` — extra columns appended after
                the data, one formula per data row. ``{row}`` is replaced with
                that row's 1-based sheet row and ``{first}`` with the first data
                row::

                    formula_columns={"total": "=B{row}*C{row}"}

                The formula text is passed through to Excel unchanged, so
                anything Excel accepts works: nested calls, ``SUMIFS``,
                cross-sheet references, and modern functions like ``XLOOKUP`` or
                ``TEXTJOIN`` (which are rewritten with the ``_xlfn.`` prefix and
                dynamic-array metadata automatically). Structure is checked —
                unbalanced parentheses or quotes raise — but function names are
                not, so ``=NOTAFUNC(A1)`` reaches the file and shows ``#NAME?``.
                There is no ``{last}``: rows are still
                streaming when these are written, so the final row is unknown;
                use ``totals_row`` for whole-column formulas.

        Raises:
            ValueError: If the sheet name is invalid (validated on save), or a
                merge range overlaps the header/data rows.
        """
        self._sheets.append((name, data))
        # Only options actually given are recorded, so the writers keep their
        # own defaults for the rest — and a CSV target only warns about options
        # that were really set.
        for option, value in {
            "column_width": column_width,
            "column_widths": column_widths,
            "column_formats": column_formats,
            "header_format": header_format,
            "dedupe_strings": dedupe_strings,
            "header_row": header_row,
            "merge_ranges": merge_ranges,
            "row_heights": row_heights,
            "row_formats": row_formats,
            "banded_rows": banded_rows,
            "autofilter": autofilter,
            "url_columns": url_columns,
            "totals_row": totals_row,
            "totals_label": totals_label,
            "totals_format": totals_format,
            "formula_columns": formula_columns,
        }.items():
            if value:
                self._per_sheet[option][name] = value
        return self

    # -- output -------------------------------------------------------------

    def _excel_only_options(self) -> List[str]:
        """Names of set options that only apply to ``.xlsx`` output.

        ``autofit`` and ``sanitize_formulas`` are left out: the first is on by
        default so it would fire on every CSV write, and the second is CSV-only.
        """
        workbook_wide = {
            "password": self._password,
            "float_format": self._float_format,
            "datetime_format": self._datetime_format,
            "index_columns": self._index_columns,
            "bold_headers": self._bold_headers,
            "freeze": self._freeze_panes,
        }
        names = [name for name, value in workbook_wide.items() if value]
        names += [
            option for option in _PER_SHEET_OPTIONS if self._per_sheet[option]
        ]
        return names

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
                ignored = self._excel_only_options()
                if ignored:
                    _warnings.warn(
                        "CSV/TSV output ignores Excel-only options: "
                        f"{', '.join(ignored)}. "
                        "The file will contain unformatted values; write to "
                        "'.xlsx' if you need them.",
                        stacklevel=2,
                    )
                write_csv(
                    data,
                    self._target,
                    delimiter=delimiter,
                    sanitize_formulas=self._sanitize_formulas,
                )
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
                **{
                    option: values[sheet_name]
                    for option, values in self._per_sheet.items()
                    if sheet_name in values
                },
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
                **{
                    option: values
                    for option, values in self._per_sheet.items()
                    if values
                },
            )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

__all__ = [
    # Class API
    "FastExcel",
    # Format API
    "Format",
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
