"""
generate_pptx.py - Automated Executive PowerPoint Report Generator
==================================================================
Generates an executive PowerPoint presentation (.pptx) formatted exactly
like 'Jul,2026 CoPQ cost trending.pptx', updating tables, KPIs, and
analytics from 'Database/CoPQ database 25.xlsx' and 'Database/CoPQ_type_analysis.xlsx'.
"""

import os
import sys
import copy
import argparse
import openpyxl
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from monthly_reporting import month_sheet_name, month_label
from build_monthly_deck import export_monthly_deck
from pipeline_common import save_with_fallback

# Ensure UTF-8 output encoding for Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE_PATH = os.path.join(BASE_DIR, "Jul,2026 CoPQ cost trending.pptx")
DEFAULT_DATABASE_PATH = os.path.join(BASE_DIR, "Database", "CoPQ database 25.xlsx")
DEFAULT_TYPE_ANALYSIS = os.path.join(BASE_DIR, "Database", "CoPQ_type_analysis.xlsx")
DEFAULT_OUTPUT_PATH = os.path.join(BASE_DIR, "Output", "CoPQ_cost_trending_Report.pptx")


def load_database_overview(db_path, report_month=None):
    """Load latest monthly overview from CoPQ database 25.xlsx.

    The monthly clone (Output/Monthly/YYYY-MM/Database/) is written by
    _write_overview() using raw COPQ_Input pivot data, which may differ from
    the official finalized figures pre-entered by hand in the source
    'Database/CoPQ database 25.xlsx'.  To guarantee the presentation always
    uses the official numbers, we prefer the *source* database when it
    contains the requested month sheet, and fall back to the clone otherwise.
    """
    # All-zero fallback: never present fabricated figures as real data.
    sites = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB", "CLG"]
    data = {
        s: {"Touch_up": 0.00, "Re_inspect": 0.00, "BC_Grade": 0.00, "Total": 0.00}
        for s in sites
    }
    vn_tot = 0.0
    indo_tot = 0.0
    clg_tot = 0.0

    # --- Choose the best database file ---
    # Prefer the canonical source database in BASE_DIR/Database/ when it already
    # contains the requested month sheet (manually entered official figures).
    # The monthly clone written by _write_overview() is derived from raw pivot
    # data and may not yet match the officially approved numbers.
    source_db = os.path.join(BASE_DIR, "Database", "CoPQ database 25.xlsx")
    requested_sheet = month_sheet_name(report_month) if report_month else None

    best_path = db_path  # default: use whatever caller supplies
    if os.path.exists(source_db) and source_db != db_path:
        try:
            _wb_check = openpyxl.load_workbook(source_db, read_only=True)
            if requested_sheet and requested_sheet in _wb_check.sheetnames:
                best_path = source_db   # source has the month sheet - prefer it
            elif not requested_sheet:
                best_path = source_db
            _wb_check.close()
        except Exception:
            pass

    if not os.path.exists(best_path):
        print("[!] WARNING: CoPQ database not found - report will contain ZEROS.")
        print(f"    Expected database at: {best_path}")
        print("    Fill in the official figures or supply a valid --database path.")
        return data, vn_tot, indo_tot, clg_tot

    try:
        wb = openpyxl.load_workbook(best_path, data_only=True)
        sheet_name = (
            requested_sheet if requested_sheet and requested_sheet in wb.sheetnames
            else ("Current" if "Current" in wb.sheetnames else wb.sheetnames[-1])
        )
        ws = wb[sheet_name]

        site_map = {}
        for r in range(3, 11):
            site_val = ws.cell(r, 1).value
            if not site_val:
                continue
            site_key = str(site_val).strip()
            tu = float(ws.cell(r, 2).value or 0)
            re_i = float(ws.cell(r, 3).value or 0)
            bc = float(ws.cell(r, 4).value or 0)
            # Official CoPQ presentation totals = Touch-up + Re-inspect + BC Grade
            tot = tu + re_i + bc
            site_map[site_key] = {"Touch_up": tu, "Re_inspect": re_i, "BC_Grade": bc, "Total": tot}

        if site_map:
            data.update(site_map)

        clg_tot = sum(data.get(s, {}).get("Total", 0) for s in ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"])
        vn_tot = sum(data.get(s, {}).get("Total", 0) for s in ["VH", "VH2", "VH3", "VH4"])
        indo_tot = sum(data.get(s, {}).get("Total", 0) for s in ["JV", "JV2", "JVB"])

    except Exception as e:
        print(f"[!] Warning loading CoPQ database: {e}")

    return data, vn_tot, indo_tot, clg_tot


def _set_text_preserve_format(text_frame, new_text, font_size_pt=None, bold=None, center=True):
    """
    Safely update a text frame's full text while preserving (or setting) run-level formatting.
    Clears to a single paragraph with a single run to prevent line-wrapping issues.
    """
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN

    tf = text_frame
    tf.word_wrap = False

    # Grab existing font properties from first run if available
    existing_size = None
    existing_bold = None
    existing_color = None
    try:
        first_run = tf.paragraphs[0].runs[0]
        existing_size = first_run.font.size
        existing_bold = first_run.font.bold
        existing_color = first_run.font.color.rgb if first_run.font.color and first_run.font.color.type else None
    except (IndexError, AttributeError):
        pass

    # Clear to a single paragraph with new text using python-pptx high-level API
    tf.text = new_text
    p = tf.paragraphs[0]
    if center:
        p.alignment = PP_ALIGN.CENTER

    if p.runs:
        run = p.runs[0]
        if font_size_pt is not None:
            run.font.size = Pt(font_size_pt)
        elif existing_size is not None:
            run.font.size = existing_size

        if bold is not None:
            run.font.bold = bold
        elif existing_bold is not None:
            run.font.bold = existing_bold

        if existing_color is not None:
            run.font.color.rgb = existing_color


def update_slide1(slide, overview_data, vn_tot, indo_tot, clg_tot):
    """Update Slide 1: Overview KPIs and Table."""
    print("   ├─ Updating Slide 1: CoPQ Overview KPIs & Matrix Table...")

    # 1. Update Master Table in Slide 1
    sites_col_map = {
        1: "VH",
        2: "VH2",
        3: "VH3",
        4: "VH4",
        5: "JV",
        6: "JV2",
        7: "JVB",
    }

    for shape in slide.shapes:
        if shape.has_table:
            t = shape.table
            # Row 0: Headers with Site Total in 2 lines: e.g. "VH \n$6,892.66"
            for col_idx, site_name in sites_col_map.items():
                if col_idx < len(t.columns):
                    tot_v = overview_data.get(site_name, {}).get("Total", 0.0)
                    cell = t.cell(0, col_idx)
                    from lxml import etree
                    nsmap_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    txBody = cell.text_frame._txBody
                    for p_elem in txBody.findall(f'{{{nsmap_a}}}p'):
                        txBody.remove(p_elem)

                    def _make_header_para(text_val):
                        """Build a paragraph element matching the template header cell format exactly."""
                        p = etree.Element(f'{{{nsmap_a}}}p')
                        pPr = etree.SubElement(p, f'{{{nsmap_a}}}pPr')
                        pPr.set('algn', 'ctr')
                        r = etree.SubElement(p, f'{{{nsmap_a}}}r')
                        rPr = etree.SubElement(r, f'{{{nsmap_a}}}rPr', lang='en-US', sz='1200')
                        # Dark text color matching template: bg2 scheme color at 100% luminance
                        solidFill = etree.SubElement(rPr, f'{{{nsmap_a}}}solidFill')
                        schemeClr = etree.SubElement(solidFill, f'{{{nsmap_a}}}schemeClr')
                        schemeClr.set('val', 'bg2')
                        lumMod = etree.SubElement(schemeClr, f'{{{nsmap_a}}}lumMod')
                        lumMod.set('val', '10000')
                        latin = etree.SubElement(rPr, f'{{{nsmap_a}}}latin')
                        latin.set('typeface', 'Arial')
                        latin.set('panose', '020B0604020202020204')
                        latin.set('pitchFamily', '34')
                        latin.set('charset', '0')
                        cs = etree.SubElement(rPr, f'{{{nsmap_a}}}cs')
                        cs.set('typeface', 'Arial')
                        cs.set('panose', '020B0604020202020204')
                        cs.set('pitchFamily', '34')
                        cs.set('charset', '0')
                        t_elem = etree.SubElement(r, f'{{{nsmap_a}}}t')
                        t_elem.text = text_val
                        return p

                    # Paragraph 1: Site name, Paragraph 2: Total cost
                    txBody.append(_make_header_para(f"{site_name} "))
                    txBody.append(_make_header_para(f"${tot_v:,.2f}"))

        # 2. Update KPI Callout Boxes (VN / CLG / ID donut labels)
        elif shape.has_text_frame:
            full_text = shape.text_frame.text.replace("\n", "").replace(" ", "")
            if shape.name == "Rectangle 86" or "26,138" in full_text or "26138" in full_text or "1,173" in full_text or "1173" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${clg_tot:,.2f}", font_size_pt=13, bold=True)
            elif shape.name == "Rectangle 69" or "13,961" in full_text or "13961" in full_text or "708,616" in full_text or "708616" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${vn_tot:,.2f}", font_size_pt=12.5, bold=True)
            elif shape.name == "Rectangle 88" or "12,176" in full_text or "12176" in full_text or "464,584" in full_text or "464584" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${indo_tot:,.2f}", font_size_pt=12.5, bold=True)


def update_slide4(slide, overview_data):
    """Update Slide 4: Touch-Up Paint site callout totals."""
    print("   ├─ Updating Slide 4: Touch-Up Paint Site Figures...")
    tu_vh = overview_data.get("VH", {}).get("Touch_up", 3070.46)
    tu_vh2 = overview_data.get("VH2", {}).get("Touch_up", 2280.95)
    tu_vh4 = overview_data.get("VH4", {}).get("Touch_up", 3608.75)
    tu_jv = overview_data.get("JV", {}).get("Touch_up", 1708.56)

    for shape in slide.shapes:
        if shape.has_text_frame:
            full_text = shape.text_frame.text.replace("\n", "").replace(" ", "")
            if "3,070" in full_text or "307046" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${tu_vh:,.2f}")
            elif "2,280" in full_text or "228095" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${tu_vh2:,.2f}")
            elif "3,608" in full_text or "360875" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${tu_vh4:,.2f}")
            elif "1,708" in full_text or "170856" in full_text:
                _set_text_preserve_format(shape.text_frame, f"${tu_jv:,.2f}")


def update_report_month_titles(prs, report_month):
    """Replace only month-bearing titles in the supplied template."""
    if not report_month:
        return
    label = month_label(report_month)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            if "Cost of Poor Quality (CoPQ)_" in text:
                _set_text_preserve_format(shape.text_frame, f"Cost of Poor Quality (CoPQ)_{label}", center=False)


def export_executive_pptx(
    template_path=DEFAULT_TEMPLATE_PATH,
    database_path=DEFAULT_DATABASE_PATH,
    type_analysis_path=DEFAULT_TYPE_ANALYSIS,
    output_path=DEFAULT_OUTPUT_PATH,
    report_month=None,
):
    print("=" * 65)
    print("  EXECUTIVE POWERPOINT REPORT GENERATOR")
    print("=" * 65)
    print(f"  Template:  {template_path}")
    print(f"  Database:  {database_path}")
    print(f"  Output:    {output_path}")
    print("=" * 65)

    if not os.path.exists(template_path):
        print(f"[-] Error: Template presentation not found: {template_path}")
        return False

    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)

    print("\n[+] [Step 1/3] Loading official database figures...")
    overview_data, vn_tot, indo_tot, clg_tot = load_database_overview(database_path, report_month)
    print(f"    • Total CoPQ (CLG): ${clg_tot:,.2f}")
    print(f"    • Vietnam Total:    ${vn_tot:,.2f}")
    print(f"    • Indonesia Total:  ${indo_tot:,.2f}")

    print("\n[+] [Step 2/3] Populating Presentation Slides...")
    prs = Presentation(template_path)
    update_report_month_titles(prs, report_month)

    if len(prs.slides) >= 1:
        update_slide1(prs.slides[0], overview_data, vn_tot, indo_tot, clg_tot)
    if len(prs.slides) >= 4:
        update_slide4(prs.slides[3], overview_data)

    print("\n[+] [Step 3/3] Saving final PowerPoint Presentation...")
    saved_path = save_with_fallback(prs, output_path)
    print("=" * 65)
    if saved_path == output_path:
        print(f"[+] SUCCESS! PowerPoint Report exported to:\n    {output_path}")
    print("=" * 65 + "\n")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Executive CoPQ PowerPoint Presentation.")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE_PATH, help="Path to template PPTX")
    parser.add_argument("--database", default=DEFAULT_DATABASE_PATH, help="Path to CoPQ database 25.xlsx")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_PATH, help="Path to output PPTX")
    parser.add_argument("--month", help="Report month in YYYY-MM format")
    args = parser.parse_args()

    export_executive_pptx(
        template_path=args.template,
        database_path=args.database,
        output_path=args.output,
        report_month=pd.Timestamp(args.month + "-01").date() if args.month else None,
    )
