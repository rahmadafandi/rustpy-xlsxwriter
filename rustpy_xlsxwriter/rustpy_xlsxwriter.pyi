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

ConditionalRule = Dict[str, Any]
"""One conditional-formatting rule; see :func:`write_worksheet` for the keys."""

ConditionalFormats = Dict[str, Union[ConditionalRule, List[ConditionalRule]]]
"""Conditional formats keyed by column name."""

Notes = Dict[str, Union[str, Dict[str, Any]]]
"""Header-cell notes, by column name."""

Images = List[Dict[str, Any]]
"""Images anchored to cells."""

Outline = Dict[str, Any]
"""Row and column grouping; see :func:`write_worksheet` for the keys."""

DataValidations = Dict[str, Dict[str, Any]]
"""Data validation rules keyed by column name."""

SheetView = Dict[str, Any]
"""Screen presentation; see :func:`write_worksheet` for the keys."""

IgnoreErrors = Union[List[str], Dict[str, str]]
"""Error indicators to suppress, by column."""

PageSetup = Dict[str, Any]
"""Page and print settings; see :func:`write_worksheet` for the keys."""

UrlColumns = Union[List[str], Dict[str, str]]
"""Link columns — a list of column names, or ``{url column: display-text column}``."""

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
    url_columns: Optional[UrlColumns] = None,
    totals_row: Optional[Dict[str, str]] = None,
    totals_label: Optional[str] = None,
    totals_format: Optional[Format] = None,
    formula_columns: Optional[Dict[str, str]] = None,
    na_rep: Optional[str] = None,
    inf_value: Optional[str] = None,
    page_setup: Optional[PageSetup] = None,
    conditional_formats: Optional[ConditionalFormats] = None,
    sheet_view: Optional[SheetView] = None,
    ignore_errors: Optional[IgnoreErrors] = None,
    data_validations: Optional[DataValidations] = None,
    outline: Optional[Outline] = None,
    notes: Optional[Notes] = None,
    images: Optional[Images] = None,
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
        url_columns: Columns whose text cells become clickable links. A list
            names them and the cell shows the URL; a dict maps each link column
            to the column holding its display text
            (``{"url": "product_name"}``). Values Excel rejects fall back to
            plain text.
        totals_row: ``{column_name: aggregate}`` written as formulas below the
            data. Valid: sum, average, count, min, max, product, stdev.
        totals_label: Text for the first column of the totals row.
        totals_format: Format applied to the whole totals row.
        formula_columns: ``{header: formula}`` appended after the data, one
            formula per row. ``{row}``/``{first}`` are substituted.
        na_rep: Text written for a missing value — ``None``, an Arrow null, or
            a float ``NaN``, which are deliberately one knob: pandas turns NaN
            in a float column into an Arrow null, so a setting that caught only
            true NaN would do nothing on the most common input there is.
            Default ``None`` leaves the cell empty, as every earlier version
            did, which makes a missing value indistinguishable from a blank.
        inf_value: Text written for ``inf``; ``-inf`` gets the same text with a
            ``-`` in front, matching Excel's own ``INF``/``-INF``.
        page_setup: Page and print settings, as one mapping — Excel has about
            twenty of them and a keyword each would double this signature.
            Keys: ``landscape``, ``paper_size``, ``margins`` (a dict of
            ``left``/``right``/``top``/``bottom``/``header``/``footer``, any
            omitted one keeping Excel's default), ``print_area``
            (``(first_row, first_col, last_row, last_col)``), ``repeat_rows``
            and ``repeat_columns`` (an index or a ``(first, last)`` pair),
            ``fit_to_pages`` (``(width, height)``; ``0`` lets that dimension
            run on), ``scale``, ``center_horizontally``, ``center_vertically``,
            ``print_gridlines``, ``print_headings``, ``first_page_number``,
            ``header`` and ``footer`` (Excel's ``&``-codes, e.g.
            ``"&RPage &P of &N"``). An unknown key raises, and so does setting
            ``scale`` together with ``fit_to_pages``, which Excel cannot honour
            at once.
        conditional_formats: Per-column conditional formatting, as
            ``{column: rule}`` or ``{column: [rule, rule]}``. A rule is a dict
            with a ``type``:

            - ``cell`` — ``criteria`` (``==``, ``!=``, ``>``, ``>=``, ``<``,
              ``<=``, ``between``, ``not_between``) plus ``value``, or ``min``
              and ``max`` for the two range criteria, and a ``format``
            - ``data_bar`` — optional ``color`` and ``bar_only``
            - ``2_color_scale`` / ``3_color_scale`` — optional ``min_color``,
              ``mid_color``, ``max_color``
            - ``text`` — ``criteria`` (``contains``, ``does_not_contain``,
              ``begins_with``, ``ends_with``), ``value``, ``format``
            - ``top`` — ``criteria`` (``top``, ``bottom``, ``top_percent``,
              ``bottom_percent``, default ``top``), ``value`` (default 10)
            - ``average`` — ``criteria`` (``above``, ``below``,
              ``equal_or_above``, ``equal_or_below``)
            - ``duplicate`` / ``unique``

            Rules cover the column's data rows only, never the header, and the
            range follows the rows actually written. An unknown column warns
            and is skipped; an unknown type or criteria raises.
        sheet_view: How the sheet presents on screen, as one mapping:
            ``tab_color``, ``gridlines`` (show them on screen), ``zoom``,
            ``right_to_left``, ``hidden``, ``selected``. Kept apart from
            ``page_setup``, which is about paper. An unknown key raises, as
            does ``hidden`` together with ``selected`` — Excel rejects a
            workbook whose active sheet is hidden.
        ignore_errors: Suppress Excel's green error triangles on a column.
            A list of column names means ``number_stored_as_text``, which is
            the reason anyone reaches for this: an ID, SKU or postcode column
            is digits stored as text on purpose. A dict maps a column to one
            error name instead — ``formula_error``, ``formula_differs``,
            ``formula_refers_to_empty_cells``, ``formula_omits_cells``,
            ``data_validation_error``, ``two_digit_text_year``,
            ``unlocked_cells_with_formula``, ``inconsistent_column_formula``.
            One name per column, never a list: Excel allows a single ignore
            rule per cell. Covers the data rows only, and an unknown column
            warns and is skipped.
        data_validations: Per-column data validation, as ``{column: rule}``.
            A rule is a dict with a ``type``:

            - ``list`` — ``values``, a list of strings; this is the dropdown,
              and the reason most people want the feature. Excel caps the
              inline list at 255 characters including separators, and a longer
              one raises rather than producing a file Excel refuses to open
            - ``whole_number`` / ``decimal`` / ``text_length`` — ``criteria``
              (``==``, ``!=``, ``>``, ``>=``, ``<``, ``<=``, ``between``,
              ``not_between``) with ``value``, or ``min`` and ``max`` for the
              two range criteria. A fraction given to ``whole_number`` or
              ``text_length`` raises instead of being truncated
            - ``custom`` — ``formula``
            - ``any`` — accepts anything, useful only to carry a message

            Any rule also takes ``input_title``, ``input_message``,
            ``error_title``, ``error_message``, ``error_style``
            (``stop``, ``warning``, ``information``), ``ignore_blank`` and
            ``show_dropdown``. Rules cover the data rows only, never the
            header. An unknown column warns and is skipped; an unknown type or
            criteria raises.
        outline: Collapsible row and column groups — the +/- brackets in
            Excel's margin — as one mapping:

            - ``rows`` — a list of ``{"from": int, "to": int}`` by 0-based
              sheet row, matching ``row_heights``, plus optional ``collapsed``
            - ``columns`` — a list of ``{"from": name, "to": name}`` by header
              name, plus optional ``collapsed``
            - ``symbols_above`` / ``symbols_to_left`` — which side the summary
              row or column sits on

            NOTE: asking for a row group takes the sheet out of constant-memory
            mode, the same trade-off as ``dedupe_strings``. The constant-memory
            row writer emits ``hidden`` but not ``outlineLevel``, so a
            collapsed group there would leave hidden rows with no bracket to
            reopen them. Column groups carry no such cost. An unknown column
            name warns and is skipped.
        notes: Notes on header cells, as ``{column: text}`` — where you say
            what a column means without widening it or adding a legend sheet.
            The value may instead be a dict with ``text`` plus any of
            ``author``, ``width``, ``height``, ``visible`` and
            ``background_color``. An unknown column warns and is skipped.
        images: Images anchored to cells, as a list of dicts. Each needs
            ``path`` or ``data`` (raw bytes, for a logo already in memory),
            and takes ``row`` and ``col`` (0-based, default 0), ``scale`` or
            the per-axis ``scale_x``/``scale_y``, ``fit_to_cell`` with
            ``keep_aspect_ratio``, ``alt_text`` and ``url``. Images are placed
            by index rather than by column name, since they float above the
            grid instead of belonging to a column. Identical images are stored
            once.

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
    url_columns: Optional[Dict[str, UrlColumns]] = None,
    totals_row: Optional[Dict[str, Dict[str, str]]] = None,
    totals_label: Optional[Dict[str, str]] = None,
    totals_format: Optional[Dict[str, Format]] = None,
    formula_columns: Optional[Dict[str, Dict[str, str]]] = None,
    na_rep: Optional[str] = None,
    inf_value: Optional[str] = None,
    page_setup: Optional[Dict[str, PageSetup]] = None,
    conditional_formats: Optional[Dict[str, ConditionalFormats]] = None,
    sheet_view: Optional[Dict[str, SheetView]] = None,
    ignore_errors: Optional[Dict[str, IgnoreErrors]] = None,
    data_validations: Optional[Dict[str, DataValidations]] = None,
    outline: Optional[Dict[str, Outline]] = None,
    notes: Optional[Dict[str, Notes]] = None,
    images: Optional[Dict[str, Images]] = None,
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
        na_rep: Text written for a missing value — ``None``, an Arrow null, or
            a float ``NaN``, which are deliberately one knob: pandas turns NaN
            in a float column into an Arrow null, so a setting that caught only
            true NaN would do nothing on the most common input there is.
            Default ``None`` leaves the cell empty, as every earlier version
            did, which makes a missing value indistinguishable from a blank.
        inf_value: Text written for ``inf``; ``-inf`` gets the same text with a
            ``-`` in front, matching Excel's own ``INF``/``-INF``.
        page_setup: Page and print settings, as one mapping — Excel has about
            twenty of them and a keyword each would double this signature.
            Keys: ``landscape``, ``paper_size``, ``margins`` (a dict of
            ``left``/``right``/``top``/``bottom``/``header``/``footer``, any
            omitted one keeping Excel's default), ``print_area``
            (``(first_row, first_col, last_row, last_col)``), ``repeat_rows``
            and ``repeat_columns`` (an index or a ``(first, last)`` pair),
            ``fit_to_pages`` (``(width, height)``; ``0`` lets that dimension
            run on), ``scale``, ``center_horizontally``, ``center_vertically``,
            ``print_gridlines``, ``print_headings``, ``first_page_number``,
            ``header`` and ``footer`` (Excel's ``&``-codes, e.g.
            ``"&RPage &P of &N"``). An unknown key raises, and so does setting
            ``scale`` together with ``fit_to_pages``, which Excel cannot honour
            at once.
        conditional_formats: Per-column conditional formatting, as
            ``{column: rule}`` or ``{column: [rule, rule]}``. A rule is a dict
            with a ``type``:

            - ``cell`` — ``criteria`` (``==``, ``!=``, ``>``, ``>=``, ``<``,
              ``<=``, ``between``, ``not_between``) plus ``value``, or ``min``
              and ``max`` for the two range criteria, and a ``format``
            - ``data_bar`` — optional ``color`` and ``bar_only``
            - ``2_color_scale`` / ``3_color_scale`` — optional ``min_color``,
              ``mid_color``, ``max_color``
            - ``text`` — ``criteria`` (``contains``, ``does_not_contain``,
              ``begins_with``, ``ends_with``), ``value``, ``format``
            - ``top`` — ``criteria`` (``top``, ``bottom``, ``top_percent``,
              ``bottom_percent``, default ``top``), ``value`` (default 10)
            - ``average`` — ``criteria`` (``above``, ``below``,
              ``equal_or_above``, ``equal_or_below``)
            - ``duplicate`` / ``unique``

            Rules cover the column's data rows only, never the header, and the
            range follows the rows actually written. An unknown column warns
            and is skipped; an unknown type or criteria raises.
        sheet_view: How the sheet presents on screen, as one mapping:
            ``tab_color``, ``gridlines`` (show them on screen), ``zoom``,
            ``right_to_left``, ``hidden``, ``selected``. Kept apart from
            ``page_setup``, which is about paper. An unknown key raises, as
            does ``hidden`` together with ``selected`` — Excel rejects a
            workbook whose active sheet is hidden.
        ignore_errors: Suppress Excel's green error triangles on a column.
            A list of column names means ``number_stored_as_text``, which is
            the reason anyone reaches for this: an ID, SKU or postcode column
            is digits stored as text on purpose. A dict maps a column to one
            error name instead — ``formula_error``, ``formula_differs``,
            ``formula_refers_to_empty_cells``, ``formula_omits_cells``,
            ``data_validation_error``, ``two_digit_text_year``,
            ``unlocked_cells_with_formula``, ``inconsistent_column_formula``.
            One name per column, never a list: Excel allows a single ignore
            rule per cell. Covers the data rows only, and an unknown column
            warns and is skipped.
        data_validations: Per-column data validation, as ``{column: rule}``.
            A rule is a dict with a ``type``:

            - ``list`` — ``values``, a list of strings; this is the dropdown,
              and the reason most people want the feature. Excel caps the
              inline list at 255 characters including separators, and a longer
              one raises rather than producing a file Excel refuses to open
            - ``whole_number`` / ``decimal`` / ``text_length`` — ``criteria``
              (``==``, ``!=``, ``>``, ``>=``, ``<``, ``<=``, ``between``,
              ``not_between``) with ``value``, or ``min`` and ``max`` for the
              two range criteria. A fraction given to ``whole_number`` or
              ``text_length`` raises instead of being truncated
            - ``custom`` — ``formula``
            - ``any`` — accepts anything, useful only to carry a message

            Any rule also takes ``input_title``, ``input_message``,
            ``error_title``, ``error_message``, ``error_style``
            (``stop``, ``warning``, ``information``), ``ignore_blank`` and
            ``show_dropdown``. Rules cover the data rows only, never the
            header. An unknown column warns and is skipped; an unknown type or
            criteria raises.
        outline: Collapsible row and column groups — the +/- brackets in
            Excel's margin — as one mapping:

            - ``rows`` — a list of ``{"from": int, "to": int}`` by 0-based
              sheet row, matching ``row_heights``, plus optional ``collapsed``
            - ``columns`` — a list of ``{"from": name, "to": name}`` by header
              name, plus optional ``collapsed``
            - ``symbols_above`` / ``symbols_to_left`` — which side the summary
              row or column sits on

            NOTE: asking for a row group takes the sheet out of constant-memory
            mode, the same trade-off as ``dedupe_strings``. The constant-memory
            row writer emits ``hidden`` but not ``outlineLevel``, so a
            collapsed group there would leave hidden rows with no bracket to
            reopen them. Column groups carry no such cost. An unknown column
            name warns and is skipped.
        notes: Notes on header cells, as ``{column: text}`` — where you say
            what a column means without widening it or adding a legend sheet.
            The value may instead be a dict with ``text`` plus any of
            ``author``, ``width``, ``height``, ``visible`` and
            ``background_color``. An unknown column warns and is skipped.
        images: Images anchored to cells, as a list of dicts. Each needs
            ``path`` or ``data`` (raw bytes, for a logo already in memory),
            and takes ``row`` and ``col`` (0-based, default 0), ``scale`` or
            the per-axis ``scale_x``/``scale_y``, ``fit_to_cell`` with
            ``keep_aspect_ratio``, ``alt_text`` and ``url``. Images are placed
            by index rather than by column name, since they float above the
            grid instead of belonging to a column. Identical images are stored
            once.

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
    na_rep: Optional[str] = None,
    inf_value: Optional[str] = None,
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
        na_rep: Text written for a missing value — ``None``, an Arrow null, or
            a float ``NaN``, which are deliberately one knob: pandas turns NaN
            in a float column into an Arrow null, so a setting that caught only
            true NaN would do nothing on the most common input there is.
            Default ``None`` leaves the cell empty, as every earlier version
            did, which makes a missing value indistinguishable from a blank.
        inf_value: Text written for ``inf``; ``-inf`` gets the same text with a
            ``-`` in front, matching Excel's own ``INF``/``-INF``.

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
