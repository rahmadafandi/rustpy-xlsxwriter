"""
RustPy-XlsxWriter Benchmark Script
====================================

Compares rustpy-xlsxwriter vs Python xlsxwriter for:
- Records (list of dicts): 500K and 1M rows
- Pandas DataFrame: 500K and 1M rows
- Polars DataFrame: 500K and 1M rows

Usage:
    python benchmark.py
"""

import os
import random
import time
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List

import numpy as np
import pandas as pd
import polars as pl
import xlsxwriter

from rustpy_xlsxwriter import FastExcel, write_csv

TMP_DIR = "/tmp/rustpy_benchmark"


# ---------------------------------------------------------------------------
# Data generators
# ---------------------------------------------------------------------------


def generate_records(count: int) -> List[Dict[str, Any]]:
    # Reuse the fixture builder to avoid a second copy of the schema.
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tests"))
    from conftest import _make_record  # type: ignore

    from faker import Faker

    random.seed(42)
    chunk_size = 10_000
    num_chunks = (count + chunk_size - 1) // chunk_size

    def _chunk(idx: int) -> List[Dict[str, Any]]:
        f = Faker()
        f.seed_instance(42 + idx)
        size = min(chunk_size, count - idx * chunk_size)
        base = [_make_record(f) for _ in range(20)]
        return (base * (size // len(base) + 1))[:size]

    with ThreadPoolExecutor() as pool:
        chunks = list(pool.map(_chunk, range(num_chunks)))
    return [r for c in chunks for r in c]


def generate_pandas_df(count: int) -> pd.DataFrame:
    np.random.seed(42)
    return pd.DataFrame(
        {
            "int_col": np.random.randint(0, 1000, count),
            "float_col": np.random.uniform(0, 100, count),
            "str_col": [f"row_{i}" for i in range(count)],
            "bool_col": np.random.choice([True, False], count),
        }
    )


def generate_polars_df(count: int) -> pl.DataFrame:
    np.random.seed(42)
    return pl.DataFrame(
        {
            "int_col": np.random.randint(0, 1000, count),
            "float_col": np.random.uniform(0, 100, count),
            "str_col": [f"row_{i}" for i in range(count)],
            "bool_col": np.random.choice([True, False], count),
        }
    )


# ---------------------------------------------------------------------------
# xlsxwriter baselines
# ---------------------------------------------------------------------------


def _xlsx_write_cell(ws, row, col, val) -> None:
    if val is None:
        return
    if isinstance(val, bool):
        ws.write_boolean(row, col, val)
    elif isinstance(val, (int, float, np.integer, np.floating)):
        ws.write_number(row, col, float(val))
    elif isinstance(val, dict):
        ws.write_string(row, col, str(val))
    else:
        ws.write(row, col, val)


def _xlsxwriter_write(path: str, headers, rows) -> None:
    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet()
    for col, h in enumerate(headers):
        ws.write(0, col, h)
    for i, row in enumerate(rows, start=1):
        for col, val in enumerate(row):
            _xlsx_write_cell(ws, i, col, val)
    wb.close()


def xlsxwriter_write_records(records: List[Dict[str, Any]], path: str) -> None:
    headers = list(records[0].keys())
    _xlsxwriter_write(path, headers, (tuple(r[h] for h in headers) for r in records))


def xlsxwriter_write_dataframe(df: pd.DataFrame, path: str) -> None:
    _xlsxwriter_write(path, list(df.columns), df.itertuples(index=False, name=None))


def xlsxwriter_write_polars(df_pl: pl.DataFrame, path: str) -> None:
    _xlsxwriter_write(path, df_pl.columns, df_pl.iter_rows())


# ---------------------------------------------------------------------------
# Benchmark runner
# ---------------------------------------------------------------------------


def bench(label: str, fn, *args) -> float:
    start = time.perf_counter()
    fn(*args)
    elapsed = time.perf_counter() - start
    return elapsed


def cleanup(path: str) -> None:
    if os.path.exists(path):
        os.remove(path)


def main() -> None:
    os.makedirs(TMP_DIR, exist_ok=True)

    results = []

    # --- Records ---
    for n in [500_000, 1_000_000]:
        label = f"{n:,}"
        print(f"Generating {label} records...")
        records = generate_records(n)

        p1 = os.path.join(TMP_DIR, f"records_{n}_rustpy.xlsx")
        p2 = os.path.join(TMP_DIR, f"records_{n}_xlsxwriter.xlsx")

        print(f"  rustpy-xlsxwriter...", end=" ", flush=True)
        t_r = bench("", lambda: FastExcel(p1, password="pw").sheet("B", records).save())
        print(f"{t_r:.2f}s")

        print(f"  xlsxwriter...", end=" ", flush=True)
        t_x = bench("", lambda: xlsxwriter_write_records(records, p2))
        print(f"{t_x:.2f}s")

        results.append(("Records", label, t_r, t_x))
        cleanup(p1)
        cleanup(p2)

    # --- Pandas DataFrame ---
    for n in [500_000, 1_000_000]:
        label = f"{n:,}"
        print(f"Generating Pandas DataFrame ({label} rows)...")
        df = generate_pandas_df(n)

        p1 = os.path.join(TMP_DIR, f"pandas_{n}_rustpy.xlsx")
        p2 = os.path.join(TMP_DIR, f"pandas_{n}_xlsxwriter.xlsx")

        print(f"  rustpy-xlsxwriter...", end=" ", flush=True)
        t_r = bench("", lambda: FastExcel(p1, autofit=False).sheet("B", df).save())
        print(f"{t_r:.2f}s")

        print(f"  xlsxwriter...", end=" ", flush=True)
        t_x = bench("", lambda: xlsxwriter_write_dataframe(df, p2))
        print(f"{t_x:.2f}s")

        results.append(("Pandas", label, t_r, t_x))
        cleanup(p1)
        cleanup(p2)

    # --- Polars DataFrame ---
    for n in [500_000, 1_000_000]:
        label = f"{n:,}"
        print(f"Generating Polars DataFrame ({label} rows)...")
        df_pl = generate_polars_df(n)

        p1 = os.path.join(TMP_DIR, f"polars_{n}_rustpy.xlsx")
        p2 = os.path.join(TMP_DIR, f"polars_{n}_xlsxwriter.xlsx")

        print(f"  rustpy-xlsxwriter...", end=" ", flush=True)
        t_r = bench("", lambda: FastExcel(p1, autofit=False).sheet("B", df_pl).save())
        print(f"{t_r:.2f}s")

        print(f"  xlsxwriter...", end=" ", flush=True)
        t_x = bench("", lambda: xlsxwriter_write_polars(df_pl, p2))
        print(f"{t_x:.2f}s")

        results.append(("Polars", label, t_r, t_x))
        cleanup(p1)
        cleanup(p2)

    # --- CSV ---
    for n in [500_000, 1_000_000]:
        label = f"{n:,}"
        print(f"CSV benchmark ({label} rows)...")

        def _gen_csv_rows():
            for i in range(n):
                yield {"id": i, "name": f"user_{i}", "score": i * 0.1, "active": i % 2 == 0}

        p1 = os.path.join(TMP_DIR, f"csv_{n}_rustpy.csv")
        p2 = os.path.join(TMP_DIR, f"csv_{n}_python.csv")

        print(f"  rustpy write_csv...", end=" ", flush=True)
        t_r = bench("", lambda: write_csv(_gen_csv_rows(), p1))
        print(f"{t_r:.2f}s")

        print(f"  python csv...", end=" ", flush=True)

        def _python_csv():
            import csv

            with open(p2, "w", newline="") as f:
                writer = csv.DictWriter(
                    f, fieldnames=["id", "name", "score", "active"]
                )
                writer.writeheader()
                for row in _gen_csv_rows():
                    writer.writerow(row)

        t_x = bench("", _python_csv)
        print(f"{t_x:.2f}s")

        results.append(("CSV", label, t_r, t_x))
        cleanup(p1)
        cleanup(p2)

    # --- Summary ---
    print()
    print("=" * 65)
    print(f"{'Type':<10} {'Rows':>10} {'RustPy':>10} {'Baseline':>12} {'Speedup':>10}")
    print("-" * 65)
    for typ, label, t_r, t_x in results:
        print(f"{typ:<10} {label:>10} {t_r:>9.2f}s {t_x:>11.2f}s {t_x/t_r:>8.1f}x")
    print("=" * 65)

    # cleanup tmp dir
    os.rmdir(TMP_DIR)


if __name__ == "__main__":
    main()
