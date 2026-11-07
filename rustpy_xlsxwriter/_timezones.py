"""Timezone-aware datetime columns, turned into their local wall clock.

Excel has no timezone, so an aware datetime is written as its own local wall
clock: 12:00 UTC converted to Asia/Jakarta is written as 19:00, which is what
the frame shows when printed. A Python ``datetime`` gets that for free, since
its hour and minute fields are the local wall clock. An Arrow timestamp
column is stored as UTC instants plus a zone name, and turning the name into
an offset takes a timezone database. The DataFrame libraries already carry
one, so the conversion happens here, with their own vectorised methods,
before the frame is handed to the Arrow export:

- polars: ``dt.replace_time_zone(None)``
- pandas: ``dt.tz_localize(None)``

Only the aware columns are rewritten. A frame without any comes back as the
same object, untouched. The Rust writer calls :func:`local_wall_clock` for
xlsx output only; CSV keeps the instant and writes its offset.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, List, Optional


def local_wall_clock(frame: Any, names: Optional[List[str]] = None) -> Any:
    """Return ``frame`` with its aware datetime columns as local wall clock.

    With ``names`` left out, the aware columns are found from the frame's
    schema or dtypes, which reads metadata only. The writer passes ``names``
    when the Arrow export shows an aware column that pass did not see: a
    pandas object column of aware ``datetime`` values.

    Returns ``None`` when ``names`` is given but ``frame`` is neither a polars
    nor a pandas DataFrame, so the caller can say why it cannot write it.
    Without ``names``, such an object is returned unchanged.
    """
    library = type(frame).__module__.split(".")[0]
    if library == "polars":
        return _polars(frame, names)
    if library == "pandas":
        return _pandas(frame, names)
    return None if names else frame


def _polars(frame: Any, names: Optional[List[str]]) -> Any:
    import polars as pl

    if names is None:
        names = [
            name
            for name, dtype in frame.schema.items()
            if isinstance(dtype, pl.Datetime) and dtype.time_zone is not None
        ]
        if not names:
            return frame
    return frame.with_columns(pl.col(names).dt.replace_time_zone(None))


def _is_aware_dtype(dtype: Any) -> bool:
    # ``DatetimeTZDtype`` carries ``tz``; an ``ArrowDtype`` timestamp keeps it
    # on its pyarrow type. Neither check needs pyarrow imported.
    if getattr(dtype, "tz", None) is not None:
        return True
    return getattr(getattr(dtype, "pyarrow_dtype", None), "tz", None) is not None


def _naive(value: Any) -> Any:
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


def _pandas(frame: Any, names: Optional[List[str]]) -> Any:
    if names is None:
        names = [name for name, dtype in frame.dtypes.items() if _is_aware_dtype(dtype)]
        if not names:
            return frame
    # A shallow copy: the caller's frame is left alone, and the columns not
    # rewritten are shared rather than copied.
    out = frame.copy(deep=False)
    for name in names:
        column = frame[name]
        if _is_aware_dtype(column.dtype):
            out[name] = column.dt.tz_localize(None)
        else:
            # An object column has no ``.dt``. Each value keeps its own wall
            # clock, exactly as a list of dicts would, which also covers a
            # column mixing several offsets.
            out[name] = column.map(_naive)
    return out
