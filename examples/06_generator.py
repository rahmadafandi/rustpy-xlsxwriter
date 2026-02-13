"""Stream large datasets with generators for memory efficiency.

Instead of loading all records into memory at once, use a generator
to yield rows one at a time. FastExcel processes them lazily.
"""

from rustpy_xlsxwriter import FastExcel


def generate_rows(n: int):
    """Yield rows one at a time — never holds all data in memory."""
    for i in range(n):
        yield {"id": i, "name": f"user_{i}", "score": i * 0.1}


# 100K rows streamed through a generator
FastExcel("output_generator.xlsx").sheet("Data", generate_rows(100_000)).save()
print("✅ output_generator.xlsx created (100K rows via generator)")
