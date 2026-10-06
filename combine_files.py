"""
combine_files.py - COPQ Data Combiner & Report Generator
=========================================================
Ingests raw COPQ ERP export files (.xls/.xlsx/.csv) from COPQ_Input/
and produces a structured Excel workbook matching the format of COPQ_Clean (4).xlsx:

Sheets produced:
    - JV, JV2, JV3, VH, VH2, VH3, VH4  -- raw cleaned per-site data (original header preserved)
    - COPQ_Clean                          -- combined master with 'Site' column prepended
"""

import os
import glob
import re
import sys
import argparse
# Allow imports from the Source and Output subfolders
_SOURCE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Source")
if _SOURCE_DIR not in sys.path:
    sys.path.insert(0, _SOURCE_DIR)

_OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Output")
if _OUTPUT_DIR not in sys.path:
    sys.path.insert(0, _OUTPUT_DIR)

from Combined_COPQ import run_copq_clean
from COPQ_Type_Detail import run_copq_type_detail
from COPQ_Pivot_table import run_copq_pivot_table
from COPQ_Database import run_copq_database

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from monthly_reporting import classify_copq_type, validate_sop_data
from pipeline_common import extract_site_name, save_with_fallback

# Ensure UTF-8 output encoding for Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

# Default directories relative to this script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT_DIR = os.path.join(BASE_DIR, "COPQ_Input")
DEFAULT_OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
DEFAULT_OUTPUT_FILE = os.path.join(DEFAULT_OUTPUT_DIR, "COPQ_Clean.xlsx")

STANDARD_COLUMNS = [
    "Date",
    "Category",
    "Type",
    "Factory",
    "Fty/Line",
    "Style Nbr",
    "Mold",
    "Model",
    "Defective Qty(Pair)",
    "FOB ($)",
    "Manpower",
    "Labor cost ($)",
    "Working hours",
    "SGS Retest Qty (Pic)",
    "SGS Test unit price/pic ($)",
    "SGS Retest failure C grade",
    "Ttl cost ($)",
    "PO",
    "Remark",
    "Remark (Re-inspection Station)",
]

# Type grouping for pivot analysis
TYPE_GROUP_MAP = {
    "B/C": "B/C",
    "BC": "B/C",
    "Reinspection": "Reinspection",
    "Re-inspection": "Reinspection",
    "Rework": "Rework",
    "Touch-up": "Touch-up",
    "Touchup": "Touch-up",
    "Bottom touch-up paint": "Touch-up",
}


def _find_copq_header_index(df):
    """Locate header row index (row with 'Date' and 'Factory'/'Fty' in it)."""
    for r in range(min(10, len(df))):
        row_vals = [str(x).strip().lower() for x in df.iloc[r].dropna()]
        if any("date" in v for v in row_vals) and any("factory" in v or "fty" in v for v in row_vals):
            return r
    return None


def _drop_spacer_columns(df, header_idx):
    """Drop empty spacer columns where header is empty/NaN and all data below is NaN."""
    header_row = df.iloc[header_idx]
    cols_to_drop = []
    for col_idx in range(df.shape[1]):
        hval = header_row.iloc[col_idx]
        if hval is None or (isinstance(hval, float) and np.isnan(hval)) or str(hval).strip() == "":
            if df.iloc[header_idx + 1:, col_idx].isna().all():
                cols_to_drop.append(df.columns[col_idx])
    return df.drop(columns=cols_to_drop) if cols_to_drop else df


def _normalize_copq_dates(df, header_idx, date_col_idx=0):
    """Normalize date column values into 'YYYY-MM-DD' strings using vectorized datetime parsing."""
    if len(df) <= header_idx + 1:
        return df
    data_series = df.iloc[header_idx + 1:, date_col_idx]
    parsed = pd.to_datetime(data_series, errors="coerce")
    formatted = parsed.dt.strftime("%Y-%m-%d")
    df.iloc[header_idx + 1:, date_col_idx] = data_series.where(parsed.isna(), formatted)
    return df


def clean_copq_dataframe(raw_df):
    """
    Clean raw COPQ dataframe exported from ERP system.
    Returns a dataframe where:
      - Row 0 is the original title row ("Cost Of Poor Quality")
      - Row 1 is the column header row
      - Rows 2+ are data rows with cleaned dates (as 'YYYY-MM-DD' strings)
    """
    df = raw_df.copy()

    # 0. Check if headers are already in df.columns (e.g. from read_html or pre-parsed DataFrames)
    col_vals = [str(c).strip().lower() for c in df.columns]
    if any("date" in v for v in col_vals) and any("factory" in v or "type" in v or "fty" in v for v in col_vals):
        header_row = pd.DataFrame([list(df.columns)], columns=range(df.shape[1]))
        title_row = pd.DataFrame([["Cost Of Poor Quality"] + [""] * (df.shape[1] - 1)], columns=range(df.shape[1]))
        data_df = df.copy()
        data_df.columns = range(df.shape[1])
        df = pd.concat([title_row, header_row, data_df], ignore_index=True)
        return _normalize_copq_dates(df, 1, date_col_idx=0)

    # 1. Locate header row index (row with "Date" and "Factory"/"Fty" in it)
    header_idx = _find_copq_header_index(df)
    if header_idx is None:
        return df

    # Drop empty spacer columns (header is None/NaN and all data below is NaN)
    df = _drop_spacer_columns(df, header_idx)

    # Reconstruct df to remove any blank spacer rows between row 0 (title) and header_idx:
    # Row 0 = Title row ("Cost Of Poor Quality")
    # Row 1 = Header row (from original header_idx)
    # Row 2+ = Data rows (from original header_idx + 1 onwards)
    if header_idx > 1:
        title_row = df.iloc[[0]]
        content_rows = df.iloc[header_idx:]
        df = pd.concat([title_row, content_rows], ignore_index=True)
        header_idx = 1
    elif header_idx == 0:
        # No separate title row in raw export; insert standard title row at row 0
        title_df = pd.DataFrame([["Cost Of Poor Quality"] + [""] * (df.shape[1] - 1)])
        df = pd.concat([title_df, df], ignore_index=True)
        header_idx = 1

    # 3. Drop footer 'Count=...' row
    if not df.empty:
        last_val = str(df.iloc[-1, 0]).strip()
        if last_val.startswith("Count=") or "total" in last_val.lower():
            df = df.iloc[:-1].reset_index(drop=True)

    # 4. Normalize date column (rows starting from header_idx + 1)
    df = _normalize_copq_dates(df, header_idx, date_col_idx=0)

    return df


def build_copq_clean_master(site_dfs: dict) -> pd.DataFrame:
    """
    Build the COPQ_Clean master sheet from all cleaned site dataframes.
    Returns a DataFrame with 'Site' as first column, proper column headers, and all data.
    """
    all_rows = []
    for site_name, df in site_dfs.items():
        # Find header row index
        header_idx = None
        for r in range(min(5, len(df))):
            row_vals = [str(x).strip().lower() for x in df.iloc[r].dropna()]
            if any("date" in v for v in row_vals):
                header_idx = r
                break
        if header_idx is None:
            continue

        headers = list(df.iloc[header_idx])
        data = df.iloc[header_idx + 1:].copy()
        data.columns = range(data.shape[1])

        # Align columns to STANDARD_COLUMNS
        num_cols = min(len(headers), len(STANDARD_COLUMNS))
        col_mapping = {}
        for i, hdr in enumerate(headers[:num_cols]):
            hdr_str = str(hdr).strip() if hdr is not None else ""
            # Find best matching standard column
            for sc in STANDARD_COLUMNS:
                if sc.lower() == hdr_str.lower():
                    col_mapping[i] = sc
                    break
            if i not in col_mapping:
                col_mapping[i] = hdr_str if hdr_str else f"Col_{i}"

        data_renamed = pd.DataFrame()
        for orig_idx, col_name in col_mapping.items():
            if orig_idx < data.shape[1]:
                data_renamed[col_name] = data.iloc[:, orig_idx].values
        data_renamed.insert(0, "Site", site_name)
        all_rows.append(data_renamed)

    if not all_rows:
        return pd.DataFrame()

    master = pd.concat(all_rows, ignore_index=True)
    if "Type" in master.columns:
        master = master[~master["Type"].astype(str).str.strip().str.lower().isin(["rework"])].reset_index(drop=True)
    return master


def build_copq_report(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build COPQ_Report sheet matching the structure in COPQ_Clean (4).xlsx:
    Rows: one per site, Columns: Factory | Touch up paint | Re-inspection | B/C Grade | Rework | Total cost ($)
    """
    sites_ordered = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JV3"]
    master_df = master_df.copy()
    # SOP: Touch-up Paint is identified from the Reinspection Station remark,
    # not only from the ERP Type label.
    master_df["_SOP_Type"] = master_df.apply(classify_copq_type, axis=1)

    report_rows = []
    for site in sites_ordered:
        site_df = master_df[master_df["Site"] == site]
        if site_df.empty:
            continue

        def sum_cost(type_vals):
            subset = site_df[site_df["_SOP_Type"].isin(type_vals)]
            if "Ttl cost ($)" in subset.columns:
                return pd.to_numeric(subset["Ttl cost ($)"], errors="coerce").sum()
            return 0.0

        tu_cost = sum_cost(["Touch-up"])
        reinsp_cost = sum_cost(["Reinspection"])
        bc_cost = sum_cost(["B/C"])
        rework_cost = sum_cost(["Rework"])
        total = tu_cost + reinsp_cost + bc_cost + rework_cost
        reinsp_tu = reinsp_cost + tu_cost

        report_rows.append({
            "Factory": site,
            "Touch up paint": round(tu_cost, 2) if tu_cost > 0 else None,
            "Re-inspection": round(reinsp_cost, 2) if reinsp_cost > 0 else None,
            "B/C Grade": round(bc_cost, 2) if bc_cost > 0 else None,
            "Rework": round(rework_cost, 2) if rework_cost > 0 else None,
            "Total cost ($)": round(total, 2) if total > 0 else None,
            "Reinspection+Bottom touch up paint": round(reinsp_tu, 2) if reinsp_tu > 0 else None,
        })

    return pd.DataFrame(report_rows)


def build_pivot_overview(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build COPQ_Pivot_Overview: B/C + Reinspection + Rework + Touch-up aggregated per site.
    """
    type_groups = {
        "B/C": ["B/C", "BC"],
        "Reinspection": ["Reinspection", "Re-inspection"],
        "Rework": ["Rework"],
        "Touch-up": ["Touch-up", "Touchup", "Bottom touch-up paint", "Touch up paint"],
    }

    sites_ordered = ["JV", "JV2", "JV3", "VH", "VH2", "VH3", "VH4"]
    pivot_rows = []
    totals = {t: {"qty": 0.0, "hrs": 0.0, "cost": 0.0} for t in type_groups}

    for site in sites_ordered:
        site_df = master_df[master_df["Site"] == site]
        if site_df.empty:
            continue
        row = {"Site": site}
        for grp_name, type_vals in type_groups.items():
            subset = site_df[site_df["Type"].isin(type_vals)]
            qty = pd.to_numeric(subset["Defective Qty(Pair)"], errors="coerce").sum() if "Defective Qty(Pair)" in subset.columns else 0.0
            hrs = pd.to_numeric(subset["Working hours"], errors="coerce").sum() if "Working hours" in subset.columns else 0.0
            cost = pd.to_numeric(subset["Ttl cost ($)"], errors="coerce").sum() if "Ttl cost ($)" in subset.columns else 0.0

            row[f"{grp_name}_Qty"] = round(qty, 2) if qty > 0 else None
            row[f"{grp_name}_Hrs"] = round(hrs, 2) if hrs > 0 else None
            row[f"{grp_name}_Cost"] = round(cost, 2) if cost > 0 else None

            totals[grp_name]["qty"] += qty
            totals[grp_name]["hrs"] += hrs
            totals[grp_name]["cost"] += cost

        pivot_rows.append(row)

    # Grand total row
    grand = {"Site": "Grand Total"}
    for grp_name in type_groups:
        grand[f"{grp_name}_Qty"] = round(totals[grp_name]["qty"], 2) if totals[grp_name]["qty"] > 0 else None
        grand[f"{grp_name}_Hrs"] = round(totals[grp_name]["hrs"], 2) if totals[grp_name]["hrs"] > 0 else None
        grand[f"{grp_name}_Cost"] = round(totals[grp_name]["cost"], 2) if totals[grp_name]["cost"] > 0 else None
    pivot_rows.append(grand)

    return pd.DataFrame(pivot_rows)


def _aggregate_copq(df, group_cols, columns):
    """Coerce the three standard numeric columns and aggregate by group_cols.
    Returns a DataFrame with Qty/Hrs/Cost sums per group."""
    out = df.copy()
    out["Defective Qty(Pair)"] = pd.to_numeric(out.get("Defective Qty(Pair)", 0), errors="coerce").fillna(0)
    out["Working hours"] = pd.to_numeric(out.get("Working hours", 0), errors="coerce").fillna(0)
    out["Ttl cost ($)"] = pd.to_numeric(out.get("Ttl cost ($)", 0), errors="coerce").fillna(0)
    return out.groupby(group_cols).agg(
        Qty=("Defective Qty(Pair)", "sum"),
        Hrs=("Working hours", "sum"),
        Cost=("Ttl cost ($)", "sum"),
    ).reset_index()[list(group_cols) + columns]


def build_pivot_top_model(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build COPQ_Pivot_TopModel: B/C aggregated by site & model.
    """
    bc_df = master_df[master_df["Type"].isin(["B/C", "BC"])].copy()
    if bc_df.empty:
        return pd.DataFrame()

    result = _aggregate_copq(bc_df, ["Site", "Model"],
                             ["Qty", "Hrs", "Cost"])
    result.columns = ["Site", "Model", "Sum of Defective Qty(Pair)", "Sum of Working hours", "Sum of Ttl cost ($)"]
    result = result.sort_values(["Site", "Sum of Ttl cost ($)"], ascending=[True, False]).reset_index(drop=True)
    return result


def build_pivot_reinsp_touchup(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build COPQ_Pivot_ReInsp&TouchUp: Reinspection & Touch-up by site & model (incl. working hours).
    """
    target_types = {
        "Reinspection": ["Reinspection", "Re-inspection"],
        "Touch-up": ["Touch-up", "Touchup", "Bottom touch-up paint", "Touch up paint"],
    }

    all_parts = []
    for type_label, type_vals in target_types.items():
        subset = master_df[master_df["Type"].isin(type_vals)].copy()
        if subset.empty:
            continue
        subset["_TypeGroup"] = type_label
        subset["Defective Qty(Pair)"] = pd.to_numeric(subset.get("Defective Qty(Pair)", 0), errors="coerce").fillna(0)
        subset["Working hours"] = pd.to_numeric(subset.get("Working hours", 0), errors="coerce").fillna(0)
        subset["Ttl cost ($)"] = pd.to_numeric(subset.get("Ttl cost ($)", 0), errors="coerce").fillna(0)
        all_parts.append(subset)

    if not all_parts:
        return pd.DataFrame()

    combined = pd.concat(all_parts)
    pivot = _aggregate_copq(combined, ["Site", "Model", "_TypeGroup"],
                            ["Qty", "Hrs", "Cost"])

    # Unstack type group into columns
    pivot_wide = pivot.pivot_table(
        index=["Site", "Model"],
        columns="_TypeGroup",
        values=["Qty", "Hrs", "Cost"],
        aggfunc="sum",
    )
    # Flatten multi-level columns: (value, TypeGroup) -> "TypeGroup_value"
    pivot_wide.columns = [f"{type_grp}_{val}" for val, type_grp in pivot_wide.columns]
    pivot_wide = pivot_wide.reset_index()
    pivot_wide = pivot_wide.sort_values(["Site", "Model"]).reset_index(drop=True)
    return pivot_wide


def build_pivot_top_factory(master_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build COPQ_Pivot_TopFactory: Reinspection aggregated by site & factory.
    """
    ri_df = master_df[master_df["Type"].isin(["Reinspection", "Re-inspection"])].copy()
    if ri_df.empty:
        return pd.DataFrame()

    result = _aggregate_copq(ri_df, ["Site", "Factory"],
                             ["Qty", "Hrs", "Cost"])
    result.columns = ["Site", "Factory", "Sum of Defective Qty(Pair)", "Sum of Working hours", "Sum of Ttl cost ($)"]
    result = result.sort_values(["Site", "Factory"]).reset_index(drop=True)
    return result


def _apply_sheet_styling(ws, header_rows=1, freeze_row=2, col_header_fill="1F4E78", col_header_font_color="FFFFFF"):
    """Apply consistent formatting to a worksheet."""
    hdr_fill = PatternFill(start_color=col_header_fill, fill_type="solid")
    hdr_font = Font(name="Calibri", size=10, bold=True, color=col_header_font_color)
    data_font = Font(name="Calibri", size=9)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws.freeze_panes = f"A{freeze_row}"
    ws.views.sheetView[0].showGridLines = True

    for r in ws.iter_rows():
        for cell in r:
            if cell.row <= header_rows:
                cell.fill = hdr_fill
                cell.font = hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin_border

    for col in ws.columns:
        max_len = 10
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)), 40))
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_len + 3


def _apply_site_sheet_styling(ws):
    """Apply styling to individual site sheets (2-row header: title + column names)."""
    # Row 1 = "Cost Of Poor Quality" title (merged-looking header)
    title_fill = PatternFill(start_color="1F4E78", fill_type="solid")
    title_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    col_hdr_fill = PatternFill(start_color="2E75B6", fill_type="solid")
    col_hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=9)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws.freeze_panes = "A3"
    ws.views.sheetView[0].showGridLines = True

    for r_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = thin_border
            if r_idx == 1:
                cell.fill = title_fill
                cell.font = title_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif r_idx == 2:
                cell.fill = col_hdr_fill
                cell.font = col_hdr_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = data_font
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    for col in ws.columns:
        max_len = 10
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, min(len(str(cell.value)), 40))
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_len + 3


def _write_pivot_title_header(ws, title, sub_columns, header_colors=None):
    """Write a 3-row header block: blank, title, blank, then sub-column headers."""
    title_fill = PatternFill(start_color="1F4E78", fill_type="solid")
    title_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    col_hdr_fill = PatternFill(start_color="2E75B6", fill_type="solid")
    col_hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Row 1: blank
    ws.append([])
    # Row 2: title
    ws.append(["", title])
    cell = ws.cell(ws.max_row, 2)
    cell.fill = title_fill
    cell.font = title_font
    # Row 3: blank
    ws.append([])
    # Row 4: column headers
    ws.append([""] + sub_columns)
    for c in range(2, 2 + len(sub_columns)):
        cell = ws.cell(ws.max_row, c)
        cell.fill = col_hdr_fill
        cell.font = col_hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.freeze_panes = f"B{ws.max_row + 1}"


def combine_copq_files(
    input_dir=DEFAULT_INPUT_DIR,
    output_file=DEFAULT_OUTPUT_FILE,
    file_paths=None,
    report_month=None,
    **kwargs,
):
    print("=" * 65)
    print("  STARTING COPQ FILE COMBINER & CLEANER")
    print("=" * 65)
    print(f"  Input Directory:  {input_dir}")
    print(f"  Output File:      {output_file}")
    if report_month:
        print(f"  Report Month:     {report_month}")
    print("=" * 65)

    if file_paths:
        all_files = [
            f for f in file_paths
            if os.path.exists(f) and not os.path.basename(f).startswith("~$")
            and os.path.splitext(f)[1].lower() in (".xlsx", ".xls", ".csv")
        ]
    else:
        if not os.path.exists(input_dir):
            print(f"[-] Input directory does not exist: {input_dir}")
            os.makedirs(input_dir, exist_ok=True)
            print(f"   Created '{input_dir}'. Please place raw COPQ files there.")
            return False

        output_dir = os.path.dirname(output_file)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        supported_extensions = ("*.csv", "*.xlsx", "*.xls")
        all_files = []
        for ext in supported_extensions:
            all_files.extend(glob.glob(os.path.join(input_dir, ext)))
        all_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]

        # Filter by report_month if specified and if files from other months exist
        if report_month and all_files:
            m_clean = report_month.strip()
            month_filtered = [f for f in all_files if m_clean in os.path.basename(f)]
            if month_filtered:
                all_files = month_filtered

    if not all_files:
        print(f"[-] No Excel or CSV files found in: {input_dir}")
        return False

    print(f"[+] Found {len(all_files)} file(s) to process:")
    for f in all_files:
        print(f"   • {os.path.basename(f)}")

    site_dfs_raw = {}    # site_name -> raw cleaned df (preserving title+header rows)
    used_sheet_names = {}

    for file_path in sorted(all_files):
        filename = os.path.basename(file_path)
        file_ext = os.path.splitext(file_path)[1].lower()
        site_name = extract_site_name(filename)

        try:
            if file_ext == ".csv":
                df = pd.read_csv(file_path, header=None, encoding="latin1")
            elif file_ext in (".xlsx", ".xls"):
                try:
                    df = pd.read_excel(file_path, header=None)
                except Exception:
                    try:
                        html_dfs = pd.read_html(file_path, header=None)
                        df = html_dfs[0] if html_dfs else pd.DataFrame()
                    except Exception:
                        raise
            else:
                continue

            cleaned_df = clean_copq_dataframe(df)

            # Deduplicate sheet names
            sheet_name = site_name[:31]
            if sheet_name in used_sheet_names:
                used_sheet_names[sheet_name] += 1
                sheet_name = f"{sheet_name[:28]}_{used_sheet_names[sheet_name]}"
            else:
                used_sheet_names[sheet_name] = 1

            site_dfs_raw[sheet_name] = cleaned_df
            data_count = max(0, len(cleaned_df) - 2)  # subtract title + header rows
            print(f"   [+] Processed sheet [{sheet_name}] ({data_count} data rows) from: {filename}")

        except Exception as e:
            print(f"   [-] Error processing {filename}: {e}")

    # ---------------------------------------------------------------
    # Build master combined dataframe
    # ---------------------------------------------------------------
    print("\n[+] Building COPQ_Clean master sheet...")
    master_df = build_copq_clean_master(site_dfs_raw)
    print(f"   ├─ Total combined records: {len(master_df)} rows")

    # ---------------------------------------------------------------
    # Write Excel Workbook
    # ---------------------------------------------------------------
    print("\n[+] Writing combined workbook...")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # Remove default empty sheet

    # === 1. Individual Site Sheets (preserving original 2-row header: title + columns) ===
    SITE_ORDER = ["JV", "JV2", "JV3", "VH", "VH2", "VH3", "VH4"]
    for site_name in SITE_ORDER:
        if site_name not in site_dfs_raw:
            continue
        df = site_dfs_raw[site_name]
        ws = wb.create_sheet(title=site_name)
        for row_vals in df.itertuples(index=False, name=None):
            ws.append(list(row_vals))
        _apply_site_sheet_styling(ws)
        print(f"   ├─ Saved sheet [{site_name}]")

    for site_name, df in site_dfs_raw.items():
        if site_name not in wb.sheetnames:
            ws = wb.create_sheet(title=site_name)
            for row_vals in df.itertuples(index=False, name=None):
                ws.append(list(row_vals))
            _apply_site_sheet_styling(ws)
            print(f"   ├─ Saved sheet [{site_name}]")

    # === 2. Execute 4-Step COPQ Automation Workflow ===
    print("\n" + "=" * 65)
    print("  EXECUTING COPQ AUTOMATION WORKFLOW")
    print("=" * 65)

    print("\n[+] [Step 1/4] Running Script COPQ Clean (Combined_COPQ)...")
    run_copq_clean(wb)

    print("\n[+] [Step 2/4] Running Script COPQ Type Detail (COPQ_Type_Detail)...")
    run_copq_type_detail(wb)

    print("\n[+] [Step 3/4] Running Script COPQ Pivot Table (COPQ_Pivot_table)...")
    run_copq_pivot_table(wb)

    print("\n[+] [Step 4/4] Running COPQ Database (COPQ_Database)...")
    run_copq_database(wb)

    save_with_fallback(wb, output_file)

    print("\n" + "=" * 65)
    print(f"[+] SUCCESS! Combined {len(site_dfs_raw)} site sheets to:")
    print(f"   {output_file}")
    print("   Sheets: " + ", ".join([ws.title for ws in wb.worksheets]))
    print("=" * 65 + "\n")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Clean and combine COPQ files into a structured Excel workbook.")
    parser.add_argument("--input", "-i", default=DEFAULT_INPUT_DIR, help="Path to COPQ_Input directory")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT_FILE, help="Path to output Excel file")
    args = parser.parse_args()

    combine_copq_files(input_dir=args.input, output_file=args.output)
