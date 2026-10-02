"""
COPQ_Database.py - Script COPQ Database
=======================================
Python implementation of Output/COPQ_Database.ts.

Workflow Step 4:
- Targets sheet 'COPQ_Clean'.
- Adds/updates column 'COPQ Category' (mapping B/C -> B/C Grade, Reinspection -> Touch-up paint / Re-inspection, Rework).
- Renames 'JVB' to 'JV3' for site uniformity.
- Creates executive report sheet 'COPQ_Report':
    - Block 1: Executive Cost Summary by Factory (VH -> JV -> CLG) with currency formatting ($#,##0.00).
    - Block 2: Quantitative breakdown sub-tables:
        * Rework lines & BC Grade quantities
        * Re-Inspection quantities & work time (JV3 -> JV2 -> JV -> VH3 -> VH2 -> VH)
        * Touch-up quantities & work time (JV3 -> JV -> VH4 -> VH2 -> VH)
"""

import os
import sys
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.table import TableColumn
from openpyxl.utils import get_column_letter

# Ensure UTF-8 stdout encoding for Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

CLEAN_SHEET_NAME = "COPQ_Clean"
REPORT_SHEET_NAME = "COPQ_Report"

F_FACTORY = "Site"
F_TYPE = "Type"
F_REMARK_STATION = "Remark (Re-inspection Station)"
F_QTY = "Defective Qty(Pair)"
F_HOURS = "Working hours"
F_COST = "Ttl cost ($)"
CATEGORY_COL = "COPQ Category"

CURRENCY_FMT = "$#,##0.00"
NUMBER_FMT = "#,##0.00"

BC_GRADE_FACTORIES = ["JV2", "JV", "VH2", "VH"]
REINSPECTION_FACTORIES = ["JV3", "JV2", "JV", "VH3", "VH2", "VH"]
TOUCHUP_FACTORIES = ["JV3", "JV", "VH4", "VH2", "VH"]


def _to_float(val):
    if val is None or val == "":
        return 0.0
    try:
        s = str(val).replace("$", "").replace(",", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _blank_if_zero(n):
    return "" if (n is None or n == 0 or n == 0.0) else n


def sort_vh_to_jv(factories):
    def get_priority(name):
        u = str(name).strip().upper()
        if u.startswith("VH"):
            return 1
        if u.startswith("JV"):
            return 2
        if u.startswith("CLG"):
            return 3
        return 4
    return sorted(factories, key=lambda a: (get_priority(a), str(a)))


def sort_jv_to_vh(factories):
    def get_priority(name):
        u = str(name).strip().upper()
        if u.startswith("JV"):
            return 1
        if u.startswith("VH"):
            return 2
        if u.startswith("CLG"):
            return 3
        return 4
    return sorted(factories, key=lambda a: (get_priority(a), str(a)))


def run_copq_database(workbook_or_path):
    """
    Execute Script COPQ Database.
    Accepts either an openpyxl.Workbook instance or a file path.
    Returns the modified openpyxl.Workbook.
    """
    is_path = isinstance(workbook_or_path, str)
    if is_path:
        wb = openpyxl.load_workbook(workbook_or_path)
    else:
        wb = workbook_or_path

    if CLEAN_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Sheet '{CLEAN_SHEET_NAME}' not found.")

    clean_ws = wb[CLEAN_SHEET_NAME]
    headers = [str(clean_ws.cell(1, c).value or "").strip() for c in range(1, clean_ws.max_column + 1)]

    def col_idx(col_name):
        try:
            return headers.index(col_name) + 1
        except ValueError:
            return -1

    factory_c = col_idx(F_FACTORY)
    type_c = col_idx(F_TYPE)
    remark_c = col_idx(F_REMARK_STATION)
    qty_c = col_idx(F_QTY)
    hours_c = col_idx(F_HOURS)
    cost_c = col_idx(F_COST)

    if remark_c == -1:
        for i, h in enumerate(headers):
            if "remark" in h.lower():
                remark_c = i + 1
                break

    if factory_c == -1 or type_c == -1:
        raise ValueError(f"Required columns ('{F_FACTORY}', '{F_TYPE}') not found in '{CLEAN_SHEET_NAME}'.")

    # Locate or create CATEGORY_COL
    if CATEGORY_COL in headers:
        cat_c = headers.index(CATEGORY_COL) + 1
    else:
        cat_c = len(headers) + 1
        headers.append(CATEGORY_COL)
        hdr_cell = clean_ws.cell(1, cat_c, CATEGORY_COL)
        hdr_cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
        hdr_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        hdr_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_font = Font(name="Calibri", size=9)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # 1. Read & Bucket Categories + write back to COPQ_Clean
    rows_data = []
    present_factories_set = set()
    present_factories_order = []

    for r in range(2, clean_ws.max_row + 1):
        raw_factory = str(clean_ws.cell(r, factory_c).value or "").strip()
        factory = "JV3" if raw_factory.upper() == "JVB" else raw_factory

        type_val = str(clean_ws.cell(r, type_c).value or "").strip()
        remark = str(clean_ws.cell(r, remark_c).value or "").strip().lower() if remark_c > 0 else ""

        category = ""
        if type_val == "B/C":
            category = "B/C Grade"
        elif type_val == "Reinspection":
            category = "Touch-up paint" if ("touch-up" in remark or "touch up" in remark) else "Re-inspection"
        elif type_val == "Rework":
            category = "Rework"

        qty_val = _to_float(clean_ws.cell(r, qty_c).value) if qty_c > 0 else 0.0
        hours_val = _to_float(clean_ws.cell(r, hours_c).value) if hours_c > 0 else 0.0
        cost_val = _to_float(clean_ws.cell(r, cost_c).value) if cost_c > 0 else 0.0

        cell = clean_ws.cell(r, cat_c, category)
        cell.font = data_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="left", vertical="center")

        if factory and factory not in present_factories_set:
            present_factories_set.add(factory)
            present_factories_order.append(factory)

        rows_data.append({
            "factory": factory,
            "type": type_val,
            "category": category,
            "qty": qty_val,
            "hours": hours_val,
            "cost": cost_val,
        })

    # Update Table in COPQ_Clean
    for t_name in list(clean_ws.tables.keys()):
        tab = clean_ws.tables[t_name]
        new_col_letter = get_column_letter(max(cat_c, clean_ws.max_column))
        tab.ref = f"A1:{new_col_letter}{clean_ws.max_row}"
        existing_col_names = [col.name for col in tab.tableColumns]
        if CATEGORY_COL not in existing_col_names:
            tab.tableColumns.append(TableColumn(id=len(tab.tableColumns) + 1, name=CATEGORY_COL))

    clean_ws.column_dimensions[get_column_letter(cat_c)].width = 18

    # 2. Aggregations
    CATS = ["Touch-up paint", "Re-inspection", "B/C Grade", "Rework"]
    cost_by_factory = {f: {c: 0.0 for c in CATS} for f in present_factories_order}
    bc_qty = {f: 0.0 for f in present_factories_order}
    reinsp_qty = {f: 0.0 for f in present_factories_order}
    reinsp_hours = {f: 0.0 for f in present_factories_order}
    touchup_qty = {f: 0.0 for f in present_factories_order}
    touchup_hours = {f: 0.0 for f in present_factories_order}

    for r in rows_data:
        f = r["factory"]
        cat = r["category"]
        if not f:
            continue
        if cat and cat in cost_by_factory.get(f, {}):
            cost_by_factory[f][cat] += r["cost"]

        if cat == "B/C Grade":
            bc_qty[f] = bc_qty.get(f, 0.0) + r["qty"]
        elif cat == "Re-inspection":
            reinsp_qty[f] = reinsp_qty.get(f, 0.0) + r["qty"]
            reinsp_hours[f] = reinsp_hours.get(f, 0.0) + r["hours"]
        elif cat == "Touch-up paint":
            touchup_qty[f] = touchup_qty.get(f, 0.0) + r["qty"]
            touchup_hours[f] = touchup_hours.get(f, 0.0) + r["hours"]

    block1_factories = sort_vh_to_jv(present_factories_order)
    rework_factories = sort_jv_to_vh(present_factories_order)

    # 3. Create Report Sheet
    if REPORT_SHEET_NAME in wb.sheetnames:
        del wb[REPORT_SHEET_NAME]
    report_ws = wb.create_sheet(title=REPORT_SHEET_NAME)

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", fill_type="solid")
    sub_header_fill = PatternFill(start_color="2E75B6", fill_type="solid")

    def write_table_block(start_r, start_c, headers, data, is_currency=False, use_secondary_header=False):
        """Write a block of headers + data cells at (start_r, start_c) [1-based]."""
        fill = sub_header_fill if use_secondary_header else header_fill
        # Header row
        for i, h in enumerate(headers):
            cell = report_ws.cell(start_r, start_c + i, h)
            cell.font = header_font
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        report_ws.row_dimensions[start_r].height = 24

        # Data rows
        for r_idx, row_values in enumerate(data):
            row_num = start_r + 1 + r_idx
            for c_idx, val in enumerate(row_values):
                col_num = start_c + c_idx
                cell = report_ws.cell(row_num, col_num, val)
                cell.font = data_font
                cell.border = thin_border
                if c_idx == 0:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if isinstance(val, (int, float)):
                        cell.number_format = CURRENCY_FMT if is_currency else NUMBER_FMT

    # --- BLOCK 1: Summary Table ---
    main_header = [
        "Factory",
        "Touch up paint",
        "Re-inspection",
        "B/C Grade",
        "Rework",
        "Total cost ($)",
        "Reinspection+Bottom touch up paint"
    ]

    main_body_data = []
    for f in block1_factories:
        m = cost_by_factory.get(f, {})
        touchup = m.get("Touch-up paint", 0.0)
        reinsp = m.get("Re-inspection", 0.0)
        bc = m.get("B/C Grade", 0.0)
        rework = ""
        total_cost = touchup + reinsp + bc
        reinsp_plus_tu = touchup + reinsp

        main_body_data.append([
            f,
            _blank_if_zero(touchup),
            _blank_if_zero(reinsp),
            _blank_if_zero(bc),
            rework,
            _blank_if_zero(total_cost),
            _blank_if_zero(reinsp_plus_tu),
        ])

    write_table_block(1, 1, main_header, main_body_data, is_currency=True)

    # --- BLOCK 2: Sub-tables (starts at row 14, 1-based) ---
    block2_start_row = 14

    # Sub-block 1: Rework (Col A-B) & BC Grade (Col D-E)
    rework_data = [[f, ""] for f in rework_factories]
    bc_data = [[f, _blank_if_zero(bc_qty.get(f, 0.0))] for f in BC_GRADE_FACTORIES]

    write_table_block(block2_start_row, 1, ["Rework", "Number of lines"], rework_data, is_currency=False, use_secondary_header=True)
    write_table_block(block2_start_row, 4, ["BC Grade", "Quantity (prs)"], bc_data, is_currency=False, use_secondary_header=True)

    # Sub-block 2: Re-Inspection
    reinsp_row = block2_start_row + len(rework_factories) + 2
    reinsp_qty_data = [[f, _blank_if_zero(reinsp_qty.get(f, 0.0))] for f in REINSPECTION_FACTORIES]
    reinsp_hrs_data = [[f, _blank_if_zero(reinsp_hours.get(f, 0.0))] for f in REINSPECTION_FACTORIES]

    write_table_block(reinsp_row, 1, ["Re-Inspection", "Quantity (prs)"], reinsp_qty_data, is_currency=False, use_secondary_header=True)
    write_table_block(reinsp_row, 4, ["Re-Inspection", "Work time (hrs)"], reinsp_hrs_data, is_currency=False, use_secondary_header=True)

    # Sub-block 3: Touch-up
    touchup_row = reinsp_row + len(REINSPECTION_FACTORIES) + 2
    touchup_qty_data = [[f, _blank_if_zero(touchup_qty.get(f, 0.0))] for f in TOUCHUP_FACTORIES]
    touchup_hrs_data = [[f, _blank_if_zero(touchup_hours.get(f, 0.0))] for f in TOUCHUP_FACTORIES]

    write_table_block(touchup_row, 1, ["Touch-up", "Quantity (prs)"], touchup_qty_data, is_currency=False, use_secondary_header=True)
    write_table_block(touchup_row, 4, ["Touch-up", "Work time (hrs)"], touchup_hrs_data, is_currency=False, use_secondary_header=True)

    # Auto-fit columns on COPQ_Report
    for col in report_ws.columns:
        max_len = 10
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)), 40))
        col_letter = get_column_letter(col[0].column)
        report_ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    print(f"[+] COPQ Database completed: Populated '{CATEGORY_COL}' and generated '{REPORT_SHEET_NAME}'.")

    if is_path:
        wb.save(workbook_or_path)
    return wb


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "COPQ_Clean.xlsx")
    if not os.path.exists(target):
        print(f"Target file not found: {target}")
        sys.exit(1)
    run_copq_database(target)
