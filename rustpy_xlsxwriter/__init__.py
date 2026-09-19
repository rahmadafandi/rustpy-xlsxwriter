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

# Re-exported unchanged: the extension resolves str and os.PathLike targets
# itself, so wrapping these in Python would only hide their signatures from
# type checkers, which read them from ``rustpy_xlsxwriter.pyi``.
from .rustpy_xlsxwriter import (
    Format,
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

#: Output formats ``save()`` can produce, mapped to their CSV delimiter.
#: ``xlsx`` has none — it does not go through the CSV writer.
_FORMATS = {"xlsx": None, "csv": ",", "tsv": "\t"}


def _detect_format(target: Any) -> str:
    """Output format implied by *target*'s file extension.

    A buffer has no extension, so it falls back to ``xlsx`` — pass
    ``output_format`` explicitly to write CSV or TSV into one.
    """
    if isinstance(target, (str, _os.PathLike)):
        name = _os.fspath(target)
        if isinstance(name, bytes):
            name = name.decode("utf-8", "replace")
        name = name.lower()
        if name.endswith(".csv"):
            return "csv"
        if name.endswith(".tsv"):
            return "tsv"
    return "xlsx"


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
        output_format: Optional[str] = None,
        password: Optional[str] = None,
        autofit: bool = True,
        sanitize_formulas: bool = False,
        bom: bool = False,
        columns: Optional[List[str]] = None,
        header: bool = True,
    ) -> None:
        """Create a new writer.

        Args:
            target: File path (``str`` or :class:`os.PathLike`, e.g.
                ``pathlib.Path``) or writable binary buffer
                (e.g. ``io.BytesIO``).
            output_format: ``"xlsx"``, ``"csv"`` or ``"tsv"``. Defaults to the
                target's file extension, and to ``"xlsx"`` for a buffer, which
                has none — so this is what writes CSV into an
                :class:`io.BytesIO`. For a delimiter other than ``,`` or tab,
                call :func:`write_csv` directly.
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
            bom: CSV/TSV only. Prefix the UTF-8 byte order mark, which is what
                makes Excel on Windows read the file as UTF-8 rather than the
                system code page. Off by default so output stays
                byte-identical.
            columns: CSV/TSV only. Select and order the output columns by name.
                An unknown name raises ``ValueError``.
            header: CSV/TSV only. Write the header row (default ``True``).
        """
        if output_format is not None and output_format not in _FORMATS:
            raise ValueError(
                f"output_format must be one of {sorted(_FORMATS)}; got {output_format!r}"
            )
        self._target = target
        self._output_format = output_format
        self._password = password
        self._autofit = autofit
        self._sanitize_formulas = sanitize_formulas
        self._bom = bom
        self._columns = columns
        self._header = header
        self._sheets: List[Tuple[str, Any]] = []
        self._float_format: Optional[str] = None
        self._datetime_format: Optional[str] = None
        self._index_columns: Optional[List[str]] = None
        self._bold_headers: bool = False
        self._na_rep: Optional[str] = None
        self._inf_value: Optional[str] = None
        self._freeze_panes: Dict[str, Dict[str, int]] = {}
        # {option: {sheet_name: value}}, filled by sheet() as options are
        # given. Keyed lazily so the option names live in one place only.
        self._per_sheet: Dict[str, Dict[str, Any]] = {}

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
        na_rep: Optional[str] = None,
        inf_value: Optional[str] = None,
    ) -> "FastExcel":
        """Set number formatting and column styling.

        Args:
            float_format: Excel number format for floats (e.g. ``"0.00"``).
            datetime_format: Excel number format for datetimes
                (default ``"yyyy-mm-ddThh:mm:ss"``).
            index_columns: Column names to render **bold**.
            bold_headers: Whether to render header row in **bold**.
            na_rep: Text written for ``NaN``. Left unset, the cell is empty —
                which is what every earlier version did, so a column of NaN
                arrives as a column of blanks that cannot be told apart from
                missing data. Applies to Excel and CSV alike.
            inf_value: Text written for ``inf``; ``-inf`` gets the same text
                with a ``-`` in front, matching Excel's ``INF``/``-INF``.
        """
        if float_format is not None:
            self._float_format = float_format
        if datetime_format is not None:
            self._datetime_format = datetime_format
        if index_columns is not None:
            self._index_columns = index_columns
        if bold_headers is not None:
            self._bold_headers = bold_headers
        if na_rep is not None:
            self._na_rep = na_rep
        if inf_value is not None:
            self._inf_value = inf_value
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
        url_columns: Optional[Union[List[str], Dict[str, str]]] = None,
        totals_row: Optional[Dict[str, str]] = None,
        totals_label: Optional[str] = None,
        totals_format: Optional["Format"] = None,
        formula_columns: Optional[Dict[str, str]] = None,
        page_setup: Optional[Dict[str, Any]] = None,
        conditional_formats: Optional[Dict[str, Any]] = None,
        sheet_view: Optional[Dict[str, Any]] = None,
        ignore_errors: Optional[Union[List[str], Dict[str, str]]] = None,
        data_validations: Optional[Dict[str, Dict[str, Any]]] = None,
        outline: Optional[Dict[str, Any]] = None,
        notes: Optional[Dict[str, Any]] = None,
        images: Optional[List[Dict[str, Any]]] = None,
        sparklines: Optional[Dict[str, Dict[str, Any]]] = None,
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
            url_columns: Columns whose text cells become clickable links.
                A list names them and each cell shows the URL —
                ``["homepage"]``. A dict maps a link column to the column
                holding its display text — ``{"url": "product_name"}`` shows
                the product name and links to the URL, which is what a report
                usually wants. Accepts what Excel accepts: ``http(s)://``,
                ``mailto:``, and ``internal:Sheet2!A1`` for a link to another
                sheet. A value Excel would reject (ordinary text, or a URL past
                its 2083-character limit) is written as plain text instead, so
                a stray non-link never aborts the export. An unknown
                display-text column warns and falls back to showing the URL.
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
            page_setup: Page and print settings as one mapping, because Excel
                has about twenty of them and a keyword each would double this
                signature. Keys: ``landscape``, ``paper_size``, ``margins``
                (a dict of ``left``/``right``/``top``/``bottom``/``header``/
                ``footer``; anything omitted keeps Excel's default),
                ``print_area`` as ``(first_row, first_col, last_row,
                last_col)``, ``repeat_rows`` and ``repeat_columns`` (an index
                or a ``(first, last)`` pair — this is what puts the header on
                every printed page), ``fit_to_pages`` as ``(width, height)``
                with ``0`` letting that dimension run on, ``scale``,
                ``center_horizontally``, ``center_vertically``,
                ``print_gridlines``, ``print_headings``, ``first_page_number``,
                and ``header``/``footer`` using Excel's ``&``-codes such as
                ``"&RPage &P of &N"``. An unknown key raises, as does setting
                ``scale`` and ``fit_to_pages`` together, which Excel cannot
                honour at once.
            conditional_formats: Per-column conditional formatting, as
                ``{column: rule}`` or ``{column: [rule, rule]}`` when a column
                needs more than one. A rule is a dict with a ``type``:
                ``cell`` (``criteria`` ``==``/``!=``/``>``/``>=``/``<``/``<=``
                with ``value``, or ``between``/``not_between`` with ``min`` and
                ``max``, plus a ``format``), ``data_bar`` (optional ``color``,
                ``bar_only``), ``2_color_scale`` and ``3_color_scale``
                (optional ``min_color``, ``mid_color``, ``max_color``),
                ``text`` (``contains``/``does_not_contain``/``begins_with``/
                ``ends_with`` with ``value`` and ``format``), ``top``
                (``top``/``bottom``/``top_percent``/``bottom_percent`` with
                ``value``, default top 10), ``average``
                (``above``/``below``/``equal_or_above``/``equal_or_below``),
                and ``duplicate``/``unique``.

                Rules cover the column's data rows only — never the header —
                and the range follows the rows actually written, so no manual
                bounds. An unknown column warns and is skipped; an unknown
                type or criteria raises.
            sheet_view: How the sheet presents on screen, as one mapping:
                ``tab_color``, ``gridlines`` (on screen), ``zoom``,
                ``right_to_left``, ``hidden``, ``selected``. Separate from
                ``page_setup``, which is about paper. An unknown key raises, as
                does ``hidden`` with ``selected`` — Excel rejects a workbook
                whose active sheet is hidden.
            ignore_errors: Suppress Excel's green error triangles on a column.
                A list of column names means ``number_stored_as_text`` — the
                reason anyone wants this, since an ID, SKU or postcode column
                is digits stored as text on purpose and Excel flags every cell
                of it. A dict maps a column to another error name instead —
                one per column, never a list, since Excel allows a single
                ignore rule per cell. Covers the data rows only; an unknown
                column warns and is skipped.
            data_validations: Per-column data validation, as
                ``{column: rule}``. A rule is a dict with a ``type``:
                ``list`` (``values``, a list of strings — the dropdown, and the
                reason most people want this; Excel caps the inline list at 255
                characters including separators and a longer one raises),
                ``whole_number``/``decimal``/``text_length`` (``criteria``
                ``==``/``!=``/``>``/``>=``/``<``/``<=`` with ``value``, or
                ``between``/``not_between`` with ``min`` and ``max``; a
                fraction given to the two integer kinds raises rather than
                being truncated), ``custom`` (``formula``), or ``any``.

                Any rule also accepts ``input_title``, ``input_message``,
                ``error_title``, ``error_message``, ``error_style``
                (``stop``, ``warning``, ``information``), ``ignore_blank`` and
                ``show_dropdown``. Rules cover the data rows only, never the
                header; an unknown column warns and is skipped.
            outline: Collapsible row and column groups — the ``+``/``-``
                brackets in Excel's margin — as one mapping. ``rows`` takes a
                list of ``{"from": int, "to": int}`` by 0-based sheet row
                (matching ``row_heights``), ``columns`` a list of
                ``{"from": name, "to": name}`` by header name, both with an
                optional ``collapsed``. ``symbols_above`` and
                ``symbols_to_left`` choose which side the summary sits on.

                NOTE: a row group takes this sheet out of constant-memory
                mode, the same trade-off as ``dedupe_strings``. The
                constant-memory row writer emits ``hidden`` but not
                ``outlineLevel``, so a collapsed group there would leave hidden
                rows with no bracket to reopen them — worse than not offering
                it. Column groups carry no such cost. An unknown column name
                warns and is skipped.
            notes: Notes on header cells, as ``{column: text}`` — where to say
                what a column means without widening it or adding a legend
                sheet. The value may instead be a dict with ``text`` plus any
                of ``author``, ``width``, ``height``, ``visible`` and
                ``background_color``. An unknown column warns and is skipped.
            images: Images anchored to cells, as a list of dicts. Each needs
                ``path`` or ``data`` (raw bytes, for a logo already in memory)
                and takes ``row``/``col`` (0-based, default 0), ``scale`` or
                the per-axis ``scale_x``/``scale_y``, ``fit_to_cell`` with
                ``keep_aspect_ratio``, ``alt_text`` and ``url``. Placed by
                index rather than by column name, since an image floats above
                the grid instead of belonging to a column. Identical images are
                stored once.
            sparklines: A one-cell chart per data row, as ``{column: rule}``.
                The column is one left empty in the records; ``from`` and
                ``to`` name the span each row summarises::

                    rows = [{"q1": 1, "q2": 5, "q3": 3, "q4": 8, "trend": None}]
                    sparklines={"trend": {"from": "q1", "to": "q4"}}

                Also takes ``type`` (``line``, ``column``, ``win_lose``),
                ``color``, ``style`` and the toggles ``high_point``,
                ``low_point``, ``first_point``, ``last_point``, ``markers``,
                ``negative_points``, ``axis`` and ``right_to_left``.

                The target column must already exist: appending one would mean
                reaching into the header assembly and column accounting that
                ``formula_columns`` uses, in both row loops, for far more cost
                than the feature is worth. An unknown column warns and is
                skipped.

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
            "page_setup": page_setup,
            "conditional_formats": conditional_formats,
            "sheet_view": sheet_view,
            "ignore_errors": ignore_errors,
            "data_validations": data_validations,
            "outline": outline,
            "notes": notes,
            "images": images,
            "sparklines": sparklines,
        }.items():
            if value:
                self._per_sheet.setdefault(option, {})[name] = value
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
        names += [option for option, values in self._per_sheet.items() if values]
        return names

    def save(self) -> None:
        """Write all sheets to the target file or buffer.

        Writes the format given as ``output_format``, or the one implied by the
        target's extension: ``.csv`` → CSV, ``.tsv`` → TSV, anything else
        (including a buffer) → Excel.

        Raises:
            ValueError: If no sheets have been added.
            OSError: If there are filesystem errors while writing.
        """
        if not self._sheets:
            raise ValueError("No sheets added. Call .sheet() before .save().")

        output_format = self._output_format or _detect_format(self._target)
        delimiter = _FORMATS[output_format]
        if delimiter is not None:
            if len(self._sheets) > 1:
                raise ValueError(
                    f"CSV/TSV output supports a single sheet; got {len(self._sheets)}."
                )
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
                bom=self._bom,
                columns=self._columns,
                header=self._header,
                na_rep=self._na_rep,
                inf_value=self._inf_value,
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
                na_rep=self._na_rep,
                inf_value=self._inf_value,
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
                na_rep=self._na_rep,
                inf_value=self._inf_value,
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
