import os
from importlib.metadata import version

import pytest
import xlsxwriter

import rustpy_xlsxwriter


def setup_module():
    """Create tmp directory if it doesn't exist"""
    if not os.path.exists("tmp"):
        os.makedirs("tmp", exist_ok=True)
    else:
        # clean up tmp directory except for .gitignore
        for file in os.listdir("tmp"):
            if file not in [".gitignore"]:
                os.remove(os.path.join("tmp", file))


def test_get_version():
    assert rustpy_xlsxwriter.get_version() == version("rustpy-xlsxwriter")


def generate_test_records(count):
    """Helper function to generate test records using multiple threads"""
    import math
    from concurrent.futures import ThreadPoolExecutor

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


def generate_test_records_with_sheet_name(count, error_sheet_name=False):
    """Helper function to generate test records with sheet name"""
    records = []
    for i in range(count):
        if error_sheet_name:
            sheet_name = f"Test[]"
        else:
            sheet_name = f"Sheet{i}"
        records.append({sheet_name: generate_test_records(count)})
    return records


@pytest.mark.parametrize("record_count", [10, 10, 10])
def test_save_records_multiple_sheets(record_count):
    """Test saving records using multiple sheets with different record counts"""
    records = generate_test_records_with_sheet_name(record_count)
    filename = f"tmp/test_{record_count}_multiple_sheets.xlsx"
    assert (
        rustpy_xlsxwriter.save_records_multiple_sheets(records, filename, "password")
        is None
    )
    assert os.path.exists(filename)


@pytest.mark.parametrize("record_count", [10, 10, 10])
def test_save_error_name_sheet_records_single_sheet(record_count):
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}_single_sheet_error.xlsx"
    sheet_name = "Test[]"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.save_records(records, filename, sheet_name, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.parametrize("record_count", [10, 10, 10])
def test_save_error_name_sheet_records_multiple_sheets(record_count):
    records = generate_test_records_with_sheet_name(record_count, True)
    filename = f"tmp/test_{record_count}_multiple_sheets_error.xlsx"
    with pytest.raises(ValueError) as e:
        rustpy_xlsxwriter.save_records_multiple_sheets(records, filename, "password")
    assert (
        str(e.value)
        == "Invalid sheet name 'Test[]'. Sheet names must be <= 31 chars and cannot contain [ ] : * ? / \\"
    )
    assert not os.path.exists(filename)


@pytest.mark.parametrize(
    "record_count",
    [
        1000,
        1000,
        1000,
    ],
)
def test_save_records_single_sheet(record_count):
    """Test saving records using single thread with different record counts"""
    records = generate_test_records(record_count)
    filename = f"tmp/test_{record_count}.xlsx"
    sheet_name = f"Test {record_count}"
    assert (
        rustpy_xlsxwriter.save_records(records, filename, sheet_name, "password")
        is None
    )
    assert os.path.exists(filename)


@pytest.mark.parametrize(
    "record_count",
    [
        1000,
        1000,
        1000,
    ],
)
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


def test_get_name():
    assert rustpy_xlsxwriter.get_name() == "rustpy-xlsxwriter"


def test_get_authors():
    assert rustpy_xlsxwriter.get_authors() == "Rahmad Afandi <rahmadafandiii@gmail.com>"


def test_get_description():
    assert (
        rustpy_xlsxwriter.get_description()
        == "Rust Python bindings for rust_xlsxwriter"
    )


def test_get_repository():
    assert (
        rustpy_xlsxwriter.get_repository()
        == "https://github.com/rahmadafandi/rustpy-xlsxwriter"
    )


def test_get_homepage():
    assert (
        rustpy_xlsxwriter.get_homepage()
        == "https://github.com/rahmadafandi/rustpy-xlsxwriter"
    )


def test_get_license():
    assert rustpy_xlsxwriter.get_license() == "MIT"


def test_validate_sheet_name():
    assert rustpy_xlsxwriter.validate_sheet_name("Test") is True
    assert rustpy_xlsxwriter.validate_sheet_name("Test[") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test]") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test:") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test*") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test?") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test/") is False
    assert rustpy_xlsxwriter.validate_sheet_name("Test\\") is False
