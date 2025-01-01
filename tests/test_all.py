import pytest
from rustpy_xlsxwriter import (
    get_version,
    save_records,
    save_records_multiple_sheets,
)
import os
import xlsxwriter


def setup_module():
    """Create tmp directory if it doesn't exist"""
    if not os.path.exists("tmp"):
        os.makedirs("tmp", exist_ok=True)


def test_get_version():
    assert get_version() == "0.4.0"


def generate_test_records(count):
    """Helper function to generate test records using multiple threads"""
    import threading
    from concurrent.futures import ThreadPoolExecutor
    import math

    def generate_chunk(start, end):
        return [{"a": str(i), "b": str(i)} for i in range(start, end)]

    # Use number of CPU cores for thread count
    num_threads = os.cpu_count()
    chunk_size = math.ceil(count / num_threads)

    with ThreadPoolExecutor(max_workers=num_threads) as executor:
        futures = []
        for i in range(0, count, chunk_size):
            end = min(i + chunk_size, count)
            futures.append(executor.submit(generate_chunk, i, end))

        records = []
        for future in futures:
            records.extend(future.result())

    return records


def generate_test_records_with_sheet_name(count):
    """Helper function to generate test records with sheet name"""
    records = []
    for i in range(count):
        records.append({f"Sheet{i}": generate_test_records(count)})
    return records


@pytest.mark.parametrize("record_count", [1000, 1000, 1000])
@pytest.mark.order(1)
def test_save_records_multiple_sheets(record_count):
    """Test saving records using multiple sheets with different record counts"""
    records = generate_test_records_with_sheet_name(record_count)
    filename = f"tmp/test_{record_count}_multiple_sheets.xlsx"
    assert save_records_multiple_sheets(records, filename, "123456") is None
    assert os.path.exists(filename)


# TODO: Add this back in when we have a better solution
# @pytest.mark.parametrize(
#     "record_count",
#     [
#         1000000,
#         1000000,
#         1000000,
#     ],
# )
# @pytest.mark.order(2)
# def test_save_records_multi_thread(record_count):
#     """Test saving records using multiple threads with different record counts"""
#     records = generate_test_records(record_count)
#     filename = f"tmp/test_{record_count}_multi_thread.xlsx"
#     sheet_name = f"Test {record_count}"
#     assert save_records_multithread(records, filename, sheet_name, "123456") is None
#     assert os.path.exists(filename)


@pytest.mark.parametrize(
    "record_count",
    [
        1000000,
        1000000,
        1000000,
    ],
)
@pytest.mark.order(3)
def test_save_records_single_sheet(record_count):
    """Test saving records using single thread with different record counts"""
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}.xlsx"
    sheet_name = f"Test {record_count}"
    assert save_records(records, filename, sheet_name, "123456") is None
    assert os.path.exists(filename)


@pytest.mark.parametrize(
    "record_count",
    [
        1000000,
        1000000,
        1000000,
    ],
)
@pytest.mark.order(4)
def test_xlsxwriter(record_count):
    """Test saving records using XlsxWriter library"""
    filename = f"tmp/test_{record_count}_xlsxwriter.xlsx"
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    records = generate_test_records(record_count)

    # Write headers
    headers = records[0].keys()
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data
    for row, record in enumerate(records, start=1):
        for col, header in enumerate(headers):
            worksheet.write(row, col, record[header])

    workbook.close()
    assert os.path.exists(filename)
