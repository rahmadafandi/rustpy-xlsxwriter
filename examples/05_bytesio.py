"""Write to an in-memory buffer instead of a file.

Useful for web frameworks (Flask, FastAPI, Django) where you want
to stream the file as an HTTP response without writing to disk.
"""

import io

from rustpy_xlsxwriter import FastExcel

buf = io.BytesIO()
FastExcel(buf).sheet("Sheet1", [{"Name": "Alice", "Age": 30}]).save()

xlsx_bytes = buf.getvalue()
print(f"✅ Generated {len(xlsx_bytes)} bytes in memory")

# Example: save buffer to file (simulating HTTP response)
with open("output_from_buffer.xlsx", "wb") as f:
    f.write(xlsx_bytes)
print("✅ output_from_buffer.xlsx written from buffer")
