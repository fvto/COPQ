import openpyxl

wb = openpyxl.load_workbook('../CoPQ database 25.xlsx', data_only=False)
ws = wb['July-2026']

blocks = [
    ("RW left", 16, 19, 2),
    ("BC right", 16, 19, 5),
    ("RE left", 23, 28, 2),
    ("RE right", 23, 28, 5),
    ("TU left", 32, 36, 2),
    ("TU right", 32, 36, 5),
]

for name, r_start, r_end, c in blocks:
    print(f"\n{name} (Col {c}):")
    for r in range(r_start, r_end + 1):
        cell = ws.cell(r, c)
        print(f"  R{r}C{c}: val={repr(cell.value)}, fmt={repr(cell.number_format)}")
