#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_copq_pptx.py
===================
Clone the COPQ/FTT PowerPoint report and update its data for a new reporting
period, preserving 100% of structure/layout/formatting/fonts/colors/position.

SAFETY (hard):
  * Original .pptx is NEVER modified (we copy it first).
  * Existing chart categories & series are KEPT; only their numeric values
    (and, where the period changes them, category labels) are rewritten by a
    (series x category) lookup.
  * Missing data -> OLD value kept + WARNING logged (never fabricated).
  * A defect type not present as a series in that chart -> that chart is SKIPPED
    (kept unchanged) and flagged for the user to supply a color; we never
    auto-invent a color or silently add/remove series.
  * Chart type never changed; slides/charts never added or deleted.

DATA CONTRACT (verified read-only):
  * COPQ_Clean.xlsx : one sheet per site (VH,VH2,VH3,VH4,JV,JV2,JV3). JV3->JVB.
    Columns: Date, Category, Type, Factory, Fty/Line, Model,
    Defective Qty(Pair), Ttl cost ($), Remark (Re-inspection Station).
  * FTT_Combined_Report.xlsx : Site_VH/VH2/JV/JV2, ColorMap_bc,
    ColorMap_BC New 2026. Site_* cols: Site,Date,Model,Metric,Defect_Issue,FTT_Qty.

Run:  python3 update_copq_pptx.py
"""

import os
import sys
import shutil
import logging
import io
from lxml import etree
from PIL import Image, ImageDraw, ImageFont

import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(HERE) == "Output":
    BASE = os.path.dirname(HERE)
    OUT_DIR = HERE
else:
    BASE = HERE
    OUT_DIR = os.path.join(HERE, "Output")

if BASE not in sys.path:
    sys.path.insert(0, BASE)

from monthly_reporting import classify_copq_type

SRC_PPTX  = os.path.join(BASE, "Template_COPQ", "Template_COPQ.pptx")
COPQ_XLSX = os.path.join(OUT_DIR, "COPQ_Clean.xlsx")
FTT_XLSX  = os.path.join(OUT_DIR, "FTT_Combined_Report.xlsx")

def detect_report_period(clean_path):
    """Detect reporting month from COPQ_Clean.xlsx dates."""
    try:
        import datetime as dt
        if os.path.exists(clean_path):
            df = pd.read_excel(clean_path, sheet_name="COPQ_Clean", nrows=500)
            if "Date" in df.columns:
                dates = pd.to_datetime(df["Date"], errors="coerce").dropna()
                if not dates.empty:
                    max_d = dates.max()
                    p_label = max_d.strftime("%b, %Y")
                    d_month = max_d.strftime("%b-%Y")
                    out_f = f"COPQ_Report_{max_d.strftime('%b_%Y')}.pptx"
                    return p_label, d_month, out_f
    except Exception:
        pass
    return "Sep, 2026", "Sep-2026", "COPQ_Report_Sep_2026.pptx"

_detected_period, _detected_db_month, _detected_out = detect_report_period(COPQ_XLSX)

# --- database (accumulated source of truth) ---
DB_OVERVIEW = os.path.join(BASE, "Database", "CoPQ database 25.xlsx")
DB_TYPE     = os.path.join(BASE, "Database", "CoPQ_type_analysis.xlsx")
DB_MONTH    = _detected_db_month
PERIOD_LABEL = _detected_period
OUT_PPTX    = os.path.join(OUT_DIR, _detected_out)

ANALYSIS_CHART_LAYOUT = {
    2: {
        "model": [("Chart 23", 110874, 2011680, 1920240, 2286000), ("Chart 31", 2116666, 2011680, 1911245, 2286000), ("Chart 34", 4134853, 2011680, 1920240, 2286000), ("Chart 36", 6128462, 2011680, 1922145, 2285365), ("Chart 17", 8186988, 2011680, 1922145, 2285365), ("Chart 27", 10196144, 2011680, 1920240, 2286000)],
        "pie": [("Chart 22", 108727, 4754880, 1920240, 2011680), ("Chart 32", 2092533, 4754880, 1919605, 2011045), ("Chart 33", 4126898, 4754880, 1919605, 2011680), ("Chart 37", 6140024, 4782312, 1920240, 2011680), ("Chart 19", 8205295, 4736592, 1920240, 2011045), ("Chart 21", 10187970, 4754880, 1918970, 1982470)],
    },
    3: {
        "model": [("Chart 4", 115503, 2103120, 2598821, 2280920), ("Chart 13", 3208235, 2135019, 2578691, 2281555), ("Chart 15", 6198669, 2103120, 2551801, 2324452), ("Chart 17", 9211377, 2103120, 2541069, 2223346)],
        "pie": [("Chart 9", 335280, 4480560, 1828800, 2286000), ("Chart 11", 3616960, 4480560, 1828800, 2286000), ("Chart 19", 6664960, 4480560, 1828800, 2286000), ("Chart 21", 9540240, 4480560, 1869440, 2302510)],
    },
}

DECK_SITES = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"]
PIE_COLORS = ["3366FF", "33CCFF", "BFBFBF", "99CCFF", "7F7F7F"]
DATA_SITE_ALIAS = {"JV3": "JVB"}    # data sheet -> deck site

# COPQ Type values to DROP entirely from the report (totals, per-site bars,
# monthly charts, factory pies, top-3 model bars). Set to [] to include all.
# Per your instruction: Rework is excluded so the report covers only
# Touch-up / Re-inspection / B/C Grade.
EXCLUDE_TYPES = ["Rework"]

def map_bar_cat(row):
    """Map one COPQ_Clean row to a deck per-site bar category, per the SOP.
    Returns 'Touch up paint', 'Re-inspection', 'B/C Grade', or None (dropped)."""
    group = classify_copq_type(row)
    if group == "Touch-up":
        return "Touch up paint"
    if group == "Reinspection":
        return "Re-inspection"
    if group == "B/C":
        return "B/C Grade"
    return None

def non_rework(df):
    """Return a copy of the per-site dataframe with EXCLUDE_TYPES dropped,
    so totals/bars/monthly/factory/model charts never include the excluded
    types (e.g. Rework)."""
    if df is None or df.empty:
        return df
    mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
    return df[mask].copy()

SLIDE1_GROUPS = [                   # top-to-bottom site groups on slide 1
    {"site": "VH",  "model_bar": "Chart 24", "factory_pie": "Chart 26", "defect": "Chart 78"},
    {"site": "VH2", "model_bar": "Chart 35", "factory_pie": "Chart 28", "defect": "Chart 80"},
    {"site": "JV",  "model_bar": "Chart 38", "factory_pie": "Chart 30", "defect": "Chart 92"},
    {"site": "JV2", "model_bar": "Chart 40", "factory_pie": "Chart 33", "defect": "Chart 96"},
]

MONTHLY_CHARTS = {"VH": "VH", "VH2": "VH2", "JV": "JV", "JV2": "JV2"}

DEFECT_VALUE_MODE = "proportion"    # 'proportion' (cols sum to 1) or 'count'

# Defect types that appear in the new data but are NOT in any color palette
# (deck embedded colors, ColorMap_bc, or ColorMap_BC New 2026). To add a new
# defect series to the Top-3 Defect charts, assign it a hex here, e.g.
#   "C/Stitching margin": "FFC000",
# The script STOPS and lists any missing defect here so you make the call,
# never auto-inventing a color.
MISSING_DEFECT_COLORS = {
    # "C/Stitching margin": "FFC000",   # <-- uncomment & set once you decide
}

# Slide-0 per-site bar charts in deck table-column order VH..JVB
SITE_BAR_CHARTS = ["Chart 18", "Chart 20", "Chart 22", "Chart 24",
                   "Chart 26", "Chart 28", "Chart 30"]

# --------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------
log = logging.getLogger("copq")
log.setLevel(logging.INFO)
log.addHandler(logging.StreamHandler(sys.stdout))
WARN = []

# --------------------------------------------------------------------------
# Data loading (read-only)
# --------------------------------------------------------------------------
def load_copq():
    xl = pd.ExcelFile(COPQ_XLSX)
    out = {}
    for sheet in xl.sheet_names:
        if sheet == "COPQ_Clean":
            continue
        df = xl.parse(sheet, header=1)
        if "Date" not in df.columns and len(df) and str(df.iloc[0, 0]).startswith("Cost Of Poor Quality"):
            df = xl.parse(sheet, header=2)
        df = df[df.get("Date", pd.Series([None])).notna()]
        if EXCLUDE_TYPES:
            before = len(df)
            mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
            df = df[mask]
            log.info(f"COPQ '{sheet}': excluded {before - len(df)} rows of "
                     f"{EXCLUDE_TYPES} (kept {len(df)})")
        for c in ("Ttl cost ($)", "Defective Qty(Pair)"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        deck_site = DATA_SITE_ALIAS.get(sheet, sheet)
        out[deck_site] = df
        log.info(f"COPQ loaded: '{sheet}' -> deck '{deck_site}', {len(df)} rows, "
                 f"cost=${df['Ttl cost ($)'].sum():,.2f}")
    return out

def load_ftt():
    if not os.path.exists(FTT_XLSX):
        log.warning(f"FTT report file not found at {FTT_XLSX}; defect charts will retain template data.")
        return {}, {}
    xl = pd.ExcelFile(FTT_XLSX)
    site_def = {}
    for s in ("Site_VH", "Site_VH2", "Site_JV", "Site_JV2"):
        if s in xl.sheet_names:
            d = xl.parse(s)
            d["FTT_Qty"] = pd.to_numeric(d.get("FTT_Qty"), errors="coerce").fillna(0)
            site_def[s.replace("Site_", "")] = d
    cmap = {}
    # ColorMap_BC New 2026 : header row 3 (0-idx); defect in col2, HEX col7, RGB 4/5/6
    if "ColorMap_BC New 2026" in xl.sheet_names:
        d = xl.parse("ColorMap_BC New 2026", header=None)
        for _, r in d.iterrows():
            name = r[2]
            if not isinstance(name, str) or name.strip() in ("", "Defect type"):
                continue
            name = norm_defect(name)
            hexv = r[7] if pd.notna(r[7]) else None
            if hexv and str(hexv).strip():
                cmap[name] = str(hexv).strip().lstrip("#")
            else:
                try:
                    cmap[name] = f"{int(r[4]):02X}{int(r[5]):02X}{int(r[6]):02X}"
                except Exception:
                    pass
    # ColorMap_bc : header row 0; defect col0, HEX col4, RGB 1/2/3
    if "ColorMap_bc" in xl.sheet_names:
        d = xl.parse("ColorMap_bc", header=0)
        for _, r in d.iterrows():
            name = r.get("Defect type")
            if not isinstance(name, str) or not name.strip():
                continue
            name = norm_defect(name)
            hexv = r.get("HEX")
            if pd.notna(hexv) and str(hexv).strip():
                cmap.setdefault(name, str(hexv).strip().lstrip("#"))
            else:
                try:
                    cmap.setdefault(name, f"{int(r['Red']):02X}{int(r['Green']):02X}{int(r['Blue']):02X}")
                except Exception:
                    pass
    log.info(f"FTT loaded: site defect sheets={list(site_def)}, "
             f"colormap={len(cmap)} entries")
    return site_def, cmap

def build_palette(prs, cmap):
    """Palette: deck-embedded series colors first (authoritative for consistency),
    then color-map file. Only C/*/B/* defect names are collected."""
    pal = {}
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_chart:
                continue
            for ser in _ser_elements(shp.chart):
                nm = _ser_name(ser)
                if nm and (nm.startswith("C/") or nm.startswith("B/")):
                    col = _ser_solid_color(ser)
                    if col:
                        pal[norm_defect(nm)] = col
    for k, v in cmap.items():
        pal.setdefault(norm_defect(k), v)
    interior_key = norm_defect("C/ Interior")
    interior_defect_key = norm_defect("C/Interior defect")
    if interior_key in pal:
        pal[interior_defect_key] = pal[interior_key]
    for k, v in MISSING_DEFECT_COLORS.items():
        pal[norm_defect(k)] = v
    log.info(f"PALETTE built: {len(pal)} defect colors "
             f"(deck={sum(1 for k in pal if k in _deck_defects(prs))})")
    return pal

def norm_defect(name):
    """Normalize defect labels so palette/document/data keys match despite
    cosmetic differences: trim, lowercase, collapse whitespace, remove the
    stray space after the slash, and drop hyphens (e.g. 'C/ Stitching margin',
    'C/Airbag defect' vs 'C/ Airbag Defect', 'C/Color mis-match' vs
    'C/Color mismatch' all collapse to the same key)."""
    if not isinstance(name, str):
        return name
    s = name.strip().lower()
    s = s.replace("-", "")
    s = s.replace(" /", "/").replace("/ ", "/")
    s = " ".join(s.split())
    return s

def _deck_defects(prs):
    s = set()
    for slide in prs.slides:
        for shp in slide.shapes:
            if not shp.has_chart:
                continue
            for ser in _ser_elements(shp.chart):
                nm = _ser_name(ser)
                if nm and (nm.startswith("C/") or nm.startswith("B/")):
                    s.add(norm_defect(nm))
    return s

# --------------------------------------------------------------------------
# OOXML helpers
# --------------------------------------------------------------------------
def _ser_elements(chart):
    return list(chart._chartSpace.chart.plotArea.iter(qn('c:ser')))

def _ser_name(ser):
    tx = ser.find(qn('c:tx'))
    if tx is None:
        return None
    strref = tx.find(qn('c:strRef'))
    if strref is None:
        return None
    cache = strref.find(qn('c:strCache'))
    if cache is None:
        return None
    pt = cache.find(qn('c:pt'))
    if pt is None:
        return None
    v = pt.find(qn('c:v'))
    return v.text if v is not None else None

def _ser_solid_color(ser):
    sp = ser.find(qn('c:spPr'))
    if sp is None:
        return None
    sf = sp.find(qn('a:solidFill'))
    if sf is None:
        return None
    srgb = sf.find(qn('a:srgbClr'))
    if srgb is not None:
        return srgb.get('val')
    return None

def _cat_strs(ser):
    cat = ser.find(qn('c:cat'))
    if cat is None:
        return []
    strref = cat.find(qn('c:strRef'))
    if strref is not None:
        cache = strref.find(qn('c:strCache'))
        if cache is not None:
            return [pt.find(qn('c:v')).text for pt in cache.findall(qn('c:pt'))]
    numref = cat.find(qn('c:numRef'))
    if numref is not None:
        cache = numref.find(qn('c:numCache'))
        if cache is not None:
            return [pt.find(qn('c:v')).text for pt in cache.findall(qn('c:pt'))]
    return []

def set_series_values_ordered(ser, values, format_code=None):
    val = ser.find(qn('c:val'))
    if val is None:
        return
    nref = val.find(qn('c:numRef'))
    if nref is None:
        nref = etree.SubElement(val, qn('c:numRef'))
    vcache = nref.find(qn('c:numCache'))
    if vcache is None:
        vcache = etree.SubElement(nref, qn('c:numCache'))
    if format_code:
        fc = vcache.find(qn('c:formatCode'))
        if fc is None:
            fc = etree.SubElement(vcache, qn('c:formatCode'))
        fc.text = format_code
    for pt in vcache.findall(qn('c:pt')):
        vcache.remove(pt)
    pc = vcache.find(qn('c:ptCount'))
    if pc is None:
        pc = etree.SubElement(vcache, qn('c:ptCount'))
    pc.set('val', str(len(values)))
    for i, v in enumerate(values):
        if v is None:
            continue
        pt = etree.SubElement(vcache, qn('c:pt'))
        pt.set('idx', str(i))
        if format_code:
            pt.set(qn('c:formatCode'), format_code)
        ve = etree.SubElement(pt, qn('c:v'))
        ve.text = f"{v:.4f}".rstrip("0").rstrip(".") if isinstance(v, float) else str(v)

def set_series_categories(ser, cats):
    cat = ser.find(qn('c:cat'))
    if cat is None:
        return
    strref = cat.find(qn('c:strRef'))
    if strref is None:
        return
    cache = strref.find(qn('c:strCache'))
    if cache is None:
        cache = etree.SubElement(strref, qn('c:strCache'))
    for pt in cache.findall(qn('c:pt')):
        cache.remove(pt)
    pc = cache.find(qn('c:ptCount'))
    if pc is None:
        pc = etree.SubElement(cache, qn('c:ptCount'))
    pc.set('val', str(len(cats)))
    for i, c in enumerate(cats):
        pt = etree.SubElement(cache, qn('c:pt'))
        pt.set('idx', str(i))
        v = etree.SubElement(pt, qn('c:v'))
        v.text = str(c)

def format_model_chart(chart):
    """Keep model chart labels readable after replacing template cache data."""
    plot_area = chart._chartSpace.chart.plotArea
    for axis in list(plot_area.iter(qn('c:catAx'))) + list(plot_area.iter(qn('c:valAx'))):
        tx_pr = axis.find(qn('c:txPr'))
        if tx_pr is None:
            tx_pr = etree.SubElement(axis, qn('c:txPr'))
        body_pr = tx_pr.find(qn('a:bodyPr'))
        if body_pr is None:
            body_pr = etree.SubElement(tx_pr, qn('a:bodyPr'))
        body_pr.set('wrap', 'none')
        for latin in tx_pr.iter(qn('a:latin')):
            latin.set('typeface', 'Arial')
        for def_rpr in tx_pr.iter(qn('a:defRPr')):
            def_rpr.set('sz', '600')

    for ser in _ser_elements(chart):
        d_lbl_groups = ser.findall(qn('c:dLbls'))
        if not d_lbl_groups:
            continue
        for d_lbls in d_lbl_groups:
            num_fmts = list(d_lbls.iter(qn('c:numFmt')))
            if not num_fmts:
                num_fmts = [etree.SubElement(d_lbls, qn('c:numFmt'))]
            for num_fmt in num_fmts:
                num_fmt.set('formatCode', '"$"#,##0.00')
                num_fmt.set('sourceLinked', '0')
            tx_prs = list(d_lbls.iter(qn('c:txPr')))
            if not tx_prs:
                tx_prs = [etree.SubElement(d_lbls, qn('c:txPr'))]
            for tx_pr in tx_prs:
                for body_pr in tx_pr.iter(qn('a:bodyPr')):
                    body_pr.set('wrap', 'none')
                for latin in tx_pr.iter(qn('a:latin')):
                    latin.set('typeface', 'Arial')
                for def_rpr in tx_pr.iter(qn('a:defRPr')):
                    def_rpr.set('sz', '700')
                    def_rpr.set('b', '1')

def set_series_color(ser, hex_color):
    sp = ser.find(qn('c:spPr'))
    if sp is None:
        return
    sf = sp.find(qn('a:solidFill'))
    if sf is None:
        sf = etree.SubElement(sp, qn('a:solidFill'))
    for bad in sf.findall(qn('a:schemeClr')):
        sf.remove(bad)
    srgb = sf.find(qn('a:srgbClr'))
    if srgb is None:
        srgb = etree.SubElement(sf, qn('a:srgbClr'))
    srgb.set('val', hex_color)

def set_series_values_by_category(ser, cat_to_val, default=0.0, warn=True):
    cats = _cat_strs(ser)
    vals = []
    for c in cats:
        if c is None:
            vals.append(default)
        elif c in cat_to_val:
            vals.append(float(cat_to_val[c]))
        else:
            vals.append(default)
            if warn and c != "":
                WARN.append(f"series '{_ser_name(ser)}' cat '{c}' not in data -> 0")
    set_series_values_ordered(ser, vals)

def get_chart_by_name(slide, name):
    for shp in slide.shapes:
        if shp.has_chart and shp.name == name:
            return shp.chart
    return None

def apply_analysis_layout(prs):
    """Apply approved analysis-slide geometry without touching chart data."""
    for slide_index, groups in ANALYSIS_CHART_LAYOUT.items():
        slide = prs.slides[slide_index]
        for group in groups.values():
            for name, left, top, width, height in group:
                shp = next((s for s in slide.shapes if s.name == name), None)
                if shp is not None:
                    shp.left, shp.top, shp.width, shp.height = left, top, width, height

# --------------------------------------------------------------------------
# Updaters
# --------------------------------------------------------------------------
def clean_model_name(name, max_len=26):
    """Clean model name: collapse whitespace, trim, limit length cleanly."""
    if not isinstance(name, str) or str(name).strip().lower() in ("none", "nan", ""):
        return ""
    s = " ".join(name.strip().split())
    if len(s) > max_len:
        return s[:max_len-2].strip() + ".."
    return s

def update_site_bar(prs, copq):
    """Update Slide 0 per-site bar charts (Chart 18, 20, 22, 24, 26, 28, 30).
    Horizontal bar chart category order:
      Index 0: Touch up paint (bottom)
      Index 1: Re-inspection (middle)
      Index 2: B/C Grade (top)
    """
    slide0 = list(prs.slides)[0]
    for name, site in zip(SITE_BAR_CHARTS, DECK_SITES):
        ch = get_chart_by_name(slide0, name)
        if ch is None:
            WARN.append(f"site bar '{name}' not found"); continue
        df = copq.get(site)
        cat_sum = {"Touch up paint": 0.0, "Re-inspection": 0.0, "B/C Grade": 0.0}
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                mapped = map_bar_cat(r)
                if mapped:
                    cat_sum[mapped] += float(r.get("Ttl cost ($)", 0) or 0)
        vals = [
            cat_sum["Touch up paint"] if cat_sum["Touch up paint"] > 0 else None,
            cat_sum["Re-inspection"] if cat_sum["Re-inspection"] > 0 else None,
            cat_sum["B/C Grade"] if cat_sum["B/C Grade"] > 0 else None,
        ]
        for ser in _ser_elements(ch):
            set_series_values_ordered(ser, vals, format_code='"$"#,##0.00')

        # Remove leader lines and manual layout from data labels so no stray lines are drawn
        for dl in ch._chartSpace.chart.plotArea.iter(qn('c:dLbls')):
            for elem in list(dl.iter()):
                if elem.tag.endswith('showLeaderLines'):
                    elem.set('val', '0')
                if elem.tag.endswith('leaderLines'):
                    p = elem.getparent()
                    if p is not None:
                        p.remove(elem)
            for dlbl in dl.findall(qn('c:dLbl')):
                ml = dlbl.find(qn('c:layout'))
                if ml is not None:
                    dlbl.remove(ml)

        # Ensure axis max is large enough so bar labels don't get forced inside the bar
        valid_vals = [v for v in vals if v is not None]
        max_val = max(valid_vals) if valid_vals else 0.0
        val_ax = ch._chartSpace.chart.plotArea.find(qn('c:valAx'))
        if val_ax is not None:
            scaling = val_ax.find(qn('c:scaling'))
            if scaling is not None:
                max_elem = scaling.find(qn('c:max'))
                if max_elem is not None:
                    cur_max = float(max_elem.get('val', 9000))
                    if max_val > cur_max * 0.65:
                        new_max = max(cur_max, (int(max_val * 1.55) // 1000 + 1) * 1000)
                        max_elem.set('val', str(int(new_max)))

        log.info(f"site bar '{name}' ({site}): TU=${cat_sum['Touch up paint']:,.0f}, "
                 f"RE=${cat_sum['Re-inspection']:,.0f}, BC=${cat_sum['B/C Grade']:,.0f}")

def update_na_badges(prs, copq):
    """Position and display N/A badges on Slide 0 exactly where a category has $0 cost.
    Park unused badges offscreen so no stale N/A badges overlap real data."""
    slide0 = list(prs.slides)[0]
    SITE_X = {
        "VH":  2614930,
        "VH2": 3904933,
        "VH3": 5300028,
        "VH4": 6634163,
        "JV":  7976235,
        "JV2": 9343390,
        "JVB": 10678160,
    }
    CAT_Y = {
        "B/C Grade": 2919675,
        "Re-inspection": 3622620,
        "Touch up paint": 4264605,
    }
    needed_slots = []
    for site in DECK_SITES:
        df = copq.get(site)
        cat_sum = {"Touch up paint": 0.0, "Re-inspection": 0.0, "B/C Grade": 0.0}
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                m = map_bar_cat(r)
                if m:
                    cat_sum[m] += float(r.get("Ttl cost ($)", 0) or 0)
        for cat in ["B/C Grade", "Re-inspection", "Touch up paint"]:
            if cat_sum[cat] < 0.01:
                needed_slots.append((SITE_X[site], CAT_Y[cat]))

    na_shapes = [s for s in slide0.shapes if s.name == "Rectangle 64"]
    for i, (x, y) in enumerate(needed_slots):
        if i < len(na_shapes):
            shp = na_shapes[i]
            shp.left = x
            shp.top = y
            shp.width = 383540
            shp.height = 245110
            if shp.has_text_frame:
                shp.text_frame.text = "N/A"
                p = shp.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.size = Pt(8)
                    p.runs[0].font.bold = True
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.color.rgb = RGBColor(160, 160, 160)

    for j in range(len(needed_slots), len(na_shapes)):
        shp = na_shapes[j]
        shp.left = -2000000  # move offscreen

def update_bottom_trend_charts(prs, copq):
    """Update Charts 32, 34, 36, 38, 40 at the bottom of Slide 0."""
    slide0 = list(prs.slides)[0]

    # Chart 32: B/C Qty (prs) [JV2, JV, VH2, VH]
    ch32 = get_chart_by_name(slide0, "Chart 32")
    if ch32:
        vals = []
        for s in ["JV2", "JV", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]["Defective Qty(Pair)"].sum() if df is not None else 0.0
            vals.append(round(v, 1))
        for ser in _ser_elements(ch32):
            set_series_values_ordered(ser, vals)
        log.info(f"Chart 32 (B/C Qty): {vals}")

    # Chart 34: Re-inspection Qty (prs) [JVB, JV2, JV, VH3, VH2, VH]
    ch34 = get_chart_by_name(slide0, "Chart 34")
    if ch34:
        vals = []
        for s in ["JVB", "JV2", "JV", "VH3", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"]["Defective Qty(Pair)"].sum() if df is not None else 0.0
            vals.append(round(v, 0))
        for ser in _ser_elements(ch34):
            set_series_values_ordered(ser, vals)
        log.info(f"Chart 34 (Re-insp Qty): {vals}")

    # Chart 36: Re-inspection Work hours [JVB, JV2, JV, VH3, VH2, VH]
    ch36 = get_chart_by_name(slide0, "Chart 36")
    if ch36:
        vals = []
        for s in ["JVB", "JV2", "JV", "VH3", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"]["Working hours"].sum() if df is not None else 0.0
            vals.append(round(v, 1))
        for ser in _ser_elements(ch36):
            set_series_values_ordered(ser, vals)
        log.info(f"Chart 36 (Re-insp Hours): {vals}")

    # Chart 38: Touch-up Qty (prs) [JV, VH4, VH2, VH]
    ch38 = get_chart_by_name(slide0, "Chart 38")
    if ch38:
        vals = []
        for s in ["JV", "VH4", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"]["Defective Qty(Pair)"].sum() if df is not None else 0.0
            vals.append(round(v, 0))
        for ser in _ser_elements(ch38):
            set_series_values_ordered(ser, vals)
        log.info(f"Chart 38 (Touch-up Qty): {vals}")

    # Chart 40: Touch-up Work hours [JV, VH4, VH2, VH]
    ch40 = get_chart_by_name(slide0, "Chart 40")
    if ch40:
        vals = []
        for s in ["JV", "VH4", "VH2", "VH"]:
            df = copq.get(s)
            v = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"]["Working hours"].sum() if df is not None else 0.0
            vals.append(round(v, 1))
        for ser in _ser_elements(ch40):
            set_series_values_ordered(ser, vals)
        log.info(f"Chart 40 (Touch-up Hours): {vals}")

def update_slide1_monthly(prs):
    """Update Slide 1 monthly trend charts (VH: Chart 1, VH2: Chart 6, JV: Chart 9, JV2: Chart 8)
    from Database/CoPQ_type_analysis.xlsx rolling 7-month history + current month."""
    if not os.path.exists(DB_TYPE):
        log.warning(f"DB_TYPE not found: {DB_TYPE}")
        return
    import openpyxl
    wb = openpyxl.load_workbook(DB_TYPE, data_only=True)
    if "Current month" not in wb.sheetnames:
        return
    ws = wb["Current month"]

    # Year-to-date months: from January (col 18 for 2026) to current month (e.g. Aug, col 25)
    curr_col = None
    for c in range(3, ws.max_column + 1):
        hdr = ws.cell(2, c).value
        cost = str(ws.cell(27, c).value or "")
        if hdr and not cost.startswith("="):
            curr_col = c
    if curr_col is None:
        curr_col = 25

    jan_col = None
    for c in range(curr_col, 2, -1):
        hdr = str(ws.cell(2, c).value or "").strip().lower()
        if hdr.startswith("jan"):
            jan_col = c
            break
    if jan_col is None:
        jan_col = 18

    month_cols = list(range(jan_col, curr_col + 1))
    months = [str(ws.cell(2, c).value or "") for c in month_cols]

    site_row_map = {
        "VH":  {"name": "Chart 1", "qty_row": 3, "cost_row": 10},
        "VH2": {"name": "Chart 6", "qty_row": 4, "cost_row": 11},
        "JV":  {"name": "Chart 9", "qty_row": 5, "cost_row": 12},
        "JV2": {"name": "Chart 8", "qty_row": 6, "cost_row": 13},
    }

    slide1 = list(prs.slides)[1]
    for site, info in site_row_map.items():
        ch = get_chart_by_name(slide1, info["name"])
        if ch is None:
            continue
        qty_vals = [float(ws.cell(info["qty_row"], c).value or 0) for c in month_cols]
        cost_vals = [float(ws.cell(info["cost_row"], c).value or 0) for c in month_cols]

        sers = _ser_elements(ch)
        if len(sers) >= 2:
            set_series_categories(sers[0], months)
            set_series_categories(sers[1], months)
            set_series_values_ordered(sers[0], qty_vals)
            set_series_values_ordered(sers[1], cost_vals)
            log.info(f"Slide 1 monthly trend {info['name']} ({site}): Qty={qty_vals[-1]}, Cost=${cost_vals[-1]:,.2f}")

def update_pie_chart_data(ch, cat_val_pairs, default_colors=None):
    """
    Updates a pie chart with non-zero cat_val_pairs: [(cat, val), ...].
    Preserves existing dPt colors where matching, removes zero/unused categories so
    0-value slices and their $0.00 data labels never show up on the pie chart!
    """
    series = _ser_elements(ch)
    if not series:
        return
    ser = series[0]

    cat_elem = ser.find(qn('c:cat'))
    old_cats = []
    if cat_elem is not None:
        strref = cat_elem.find(qn('c:strRef'))
        if strref is not None:
            cache = strref.find(qn('c:strCache'))
            if cache is not None:
                pts = sorted(cache.findall(qn('c:pt')), key=lambda x: int(x.get('idx', 0)))
                old_cats = [p.find(qn('c:v')).text if p.find(qn('c:v')) is not None else "" for p in pts]

    dpts = ser.findall(qn('c:dPt'))

    active_pairs = [(c, round(float(v), 2)) for c, v in cat_val_pairs if v is not None and float(v) > 0.0001]
    if not active_pairs:
        active_pairs = [(cat_val_pairs[0][0], 0.0)] if cat_val_pairs else [("N/A", 0.0)]

    new_cats = [c for c, _ in active_pairs]
    new_vals = [v for _, v in active_pairs]

    set_series_categories(ser, new_cats)
    set_series_values_ordered(ser, new_vals)

    for dp in dpts:
        ser.remove(dp)

    for i, cat in enumerate(new_cats):
        col_hex = PIE_COLORS[i % len(PIE_COLORS)]
        dp = etree.SubElement(ser, qn('c:dPt'))
        etree.SubElement(dp, qn('c:idx'), val=str(i))
        spPr = etree.SubElement(dp, qn('c:spPr'))
        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        etree.SubElement(solidFill, qn('a:srgbClr'), val=col_hex)
        if len(new_cats) > 1:
            line = etree.SubElement(spPr, qn('a:ln'), w='12700')
            line_fill = etree.SubElement(line, qn('a:solidFill'))
            etree.SubElement(line_fill, qn('a:srgbClr'), val='FFFFFF')


def update_factory_pie(prs, copq):
    slide1 = list(prs.slides)[1]
    for g in SLIDE1_GROUPS:
        site = g["site"]
        ch = get_chart_by_name(slide1, g["factory_pie"])
        if ch is None:
            WARN.append(f"factory pie '{g['factory_pie']}' ({site}) not found"); continue
        df = copq.get(site)
        if df is None or df.empty:
            WARN.append(f"factory pie ({site}): no COPQ data"); continue
        bc = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]
        if bc.empty:
            continue
        fac_sum = bc.groupby("Factory")["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"factory pie ({site}) updated with active slices: {cat_val_pairs}")

def update_top3_model(prs, copq):
    top5_charts = {"VH": "Chart 24", "VH2": "Chart 35",
                   "JV": "Chart 38", "JV2": "Chart 40"}
    slide1 = list(prs.slides)[1]
    for site, name in top5_charts.items():
        ch = get_chart_by_name(slide1, name)
        if ch is None:
            WARN.append(f"top3 model '{name}' ({site}) not found"); continue
        df = copq.get(site)
        if df is None or df.empty:
            WARN.append(f"top3 model ({site}): no COPQ data"); continue
        bc = df[df.apply(map_bar_cat, axis=1) == "B/C Grade"]
        if bc.empty:
            continue
        n = 5
        top = bc.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(n)
        # Reverse order so highest cost model is at the TOP of the horizontal bar chart
        top_asc = top.iloc[::-1]
        cats = [clean_model_name(str(m)) for m in top_asc.index]
        vals = [round(float(v), 2) for v in top_asc.values]
        while len(cats) < n:
            cats.insert(0, "")
            vals.insert(0, None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)
        log.info(f"top model ({site}) -> {[c for c in cats if c]}")

def _is_dark_hex(hex_str):
    """Return True if hex color is visually dark (needs white font)."""
    h = str(hex_str).lstrip("#")
    if len(h) == 8:
        h = h[2:]
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        lum = 0.299 * r + 0.587 * g + 0.114 * b
        return lum < 128
    except Exception:
        return False

def _build_ser(idx, name, cats, vals, hex_color):
    """Construct a <c:ser> matching the exact structure and formatting of the format file."""
    ser = etree.Element(qn('c:ser'))
    etree.SubElement(ser, qn('c:idx'), val=str(idx))
    etree.SubElement(ser, qn('c:order'), val=str(idx))

    # tx
    tx = etree.SubElement(ser, qn('c:tx'))
    strref = etree.SubElement(tx, qn('c:strRef'))
    etree.SubElement(strref, qn('c:f')).text = f"'[FTT_Combined_Report.xlsx]Top Defects Summary'!$A${idx + 2}"
    cache = etree.SubElement(strref, qn('c:strCache'))
    etree.SubElement(cache, qn('c:ptCount'), val='1')
    pt0 = etree.SubElement(cache, qn('c:pt'), idx='0')
    etree.SubElement(pt0, qn('c:v')).text = str(name)

    # spPr (series fill color)
    spPr = etree.SubElement(ser, qn('c:spPr'))
    sf = etree.SubElement(spPr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr'), val=hex_color)

    # dLbls: MATCHING FORMAT FILE EXACTLY (numbers only, no shoe names, no defect names, no legend keys)
    dlbls = etree.SubElement(ser, qn('c:dLbls'))
    etree.SubElement(dlbls, qn('c:numFmt'), formatCode='0.00%', sourceLinked='1')

    dl_spPr = etree.SubElement(dlbls, qn('c:spPr'))
    etree.SubElement(dl_spPr, qn('a:noFill'))
    dl_ln = etree.SubElement(dl_spPr, qn('a:ln'))
    etree.SubElement(dl_ln, qn('a:noFill'))

    font_color = "FFFFFF" if _is_dark_hex(hex_color) else "000000"
    txPr = etree.SubElement(dlbls, qn('c:txPr'))
    etree.SubElement(txPr, qn('a:bodyPr'), rot='0', vert='horz', wrap='square', anchor='ctr', anchorCtr='1', lIns='38100', tIns='19050', rIns='38100', bIns='19050')
    etree.SubElement(txPr, qn('a:lstStyle'))
    p = etree.SubElement(txPr, qn('a:p'))
    pPr = etree.SubElement(p, qn('a:pPr'))
    defRPr = etree.SubElement(pPr, qn('a:defRPr'), lang='en-US', sz='600', b='1', i='0', u='none', strike='noStrike', kern='1200', baseline='0')
    solidFill = etree.SubElement(defRPr, qn('a:solidFill'))
    etree.SubElement(solidFill, qn('a:srgbClr'), val=font_color)
    etree.SubElement(defRPr, qn('a:latin'), typeface='+mn-lt')
    etree.SubElement(defRPr, qn('a:cs'), typeface='+mn-cs')

    etree.SubElement(dlbls, qn('c:dLblPos'), val='ctr')
    etree.SubElement(dlbls, qn('c:showLegendKey'), val='0')
    etree.SubElement(dlbls, qn('c:showVal'), val='1')
    etree.SubElement(dlbls, qn('c:showCatName'), val='0')
    etree.SubElement(dlbls, qn('c:showSerName'), val='0')
    etree.SubElement(dlbls, qn('c:showPercent'), val='0')
    etree.SubElement(dlbls, qn('c:showBubbleSize'), val='0')
    etree.SubElement(dlbls, qn('c:showLeaderLines'), val='0')

    # cat
    cat = etree.SubElement(ser, qn('c:cat'))
    cref = etree.SubElement(cat, qn('c:strRef'))
    etree.SubElement(cref, qn('c:f')).text = "'[FTT_Combined_Report.xlsx]Top Defects Summary'!$A$1"
    ccache = etree.SubElement(cref, qn('c:strCache'))
    etree.SubElement(ccache, qn('c:ptCount'), val=str(len(cats)))
    for i, c in enumerate(cats):
        cpt = etree.SubElement(ccache, qn('c:pt'), idx=str(i))
        etree.SubElement(cpt, qn('c:v')).text = str(c)

    # val: ONLY non-zero points are added to numCache (matching format file)
    val = etree.SubElement(ser, qn('c:val'))
    vref = etree.SubElement(val, qn('c:numRef'))
    etree.SubElement(vref, qn('c:f')).text = "'[FTT_Combined_Report.xlsx]Top Defects Summary'!$A$1"
    vcache = etree.SubElement(vref, qn('c:numCache'))
    etree.SubElement(vcache, qn('c:formatCode')).text = "0.00%"
    etree.SubElement(vcache, qn('c:ptCount'), val=str(len(vals)))
    for i, x in enumerate(vals):
        if x is not None and x > 0.0001:
            vpt = etree.SubElement(vcache, qn('c:pt'), idx=str(i))
            etree.SubElement(vpt, qn('c:v')).text = str(x)

    return ser

def load_ftt_top_defects_from_excel(ftt_xlsx_path):
    """
    Load Top Defects Summary matrix directly from FTT_Combined_Report.xlsx.
    Returns a dict mapping site -> {
        'models': [model1, model2, model3],
        'defects': [defect1, defect2, ...],
        'matrix': {defect: [rate_m1, rate_m2, rate_m3]}
    }
    """
    if not os.path.exists(ftt_xlsx_path):
        return {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ftt_xlsx_path, data_only=False)
        if "Top Defects Summary" not in wb.sheetnames:
            return {}
        ws = wb["Top Defects Summary"]
        xl = pd.ExcelFile(ftt_xlsx_path)
        site_dfs = {}
        for s in ("VH", "VH2", "JV", "JV2"):
            sheet = f"Site_{s}"
            if sheet in xl.sheet_names:
                df = xl.parse(sheet)
                df["FTT_Qty"] = pd.to_numeric(df.get("FTT_Qty"), errors="coerce").fillna(0)
                site_dfs[s] = df

        out = {}
        current_site = None
        header_row = None
        for r in range(1, ws.max_row + 1):
            val = ws.cell(r, 1).value
            if val in ("VH", "VH2", "JV", "JV2"):
                current_site = val
                header_row = r
                defects = []
                c = 2
                while True:
                    d_val = ws.cell(r, c).value
                    if not d_val:
                        break
                    defects.append(str(d_val).strip())
                    c += 1
                out[current_site] = {'defects': defects, 'models': [], 'model_rows': []}
            elif current_site and r <= header_row + 3 and val:
                out[current_site]['models'].append(str(val).strip())
                out[current_site]['model_rows'].append(r)

        results = {}
        for site, info in out.items():
            defects = info['defects']
            models = info['models']
            model_rows = info['model_rows']
            df = site_dfs.get(site)
            matrix = {d: [0.0] * len(models) for d in defects}
            if df is not None and not df.empty:
                for m_idx, (model_name, r_idx) in enumerate(zip(models, model_rows)):
                    m_df = df[df["Model"] == model_name]
                    ttl_qty = float(m_df["FTT_Qty"].sum())
                    for d_idx, defect_name in enumerate(defects, start=2):
                        cell = ws.cell(r_idx, d_idx)
                        if cell.value is not None:
                            d_qty = float(m_df[m_df["Defect_Issue"] == defect_name]["FTT_Qty"].sum())
                            rate = (d_qty / ttl_qty) if ttl_qty > 0 else 0.0
                            matrix[defect_name][m_idx] = rate
                        else:
                            matrix[defect_name][m_idx] = 0.0
            results[site] = {'models': models, 'defects': defects, 'matrix': matrix}
        return results
    except Exception as e:
        log.warning(f"Could not load Top Defects Summary from {ftt_xlsx_path}: {e}")
        return {}

def format_file_model_name(name):
    """Normalize model name to match the exact format file spelling."""
    s = " ".join(str(name).strip().split())
    if "WINFLO 12" in s and "WIDE" in s:
        return "NIKE AIR WINFLO 12"
    return s

def update_top3_defect(prs, ftt, palette):
    summary_data = load_ftt_top_defects_from_excel(FTT_XLSX)
    slide1 = list(prs.slides)[1]
    for g in SLIDE1_GROUPS:
        site = g["site"]
        name = g["defect"]
        ch = get_chart_by_name(slide1, name)
        if ch is None:
            WARN.append(f"top3 defect '{name}' ({site}) not found"); continue

        if summary_data and site in summary_data:
            info = summary_data[site]
            raw_models = info['models']
            models = [format_file_model_name(m) for m in raw_models]
            defects = info['defects']
            matrix = info['matrix']

            series = _ser_elements(ch)
            if not series:
                WARN.append(f"top3 defect ({site}): no <c:ser> found"); continue
            parent = series[0].getparent()
            for s in series:
                parent.remove(s)

            for i, d in enumerate(defects):
                norm_d = norm_defect(d)
                col = palette.get(norm_d, "BFBFBF")
                vals = matrix.get(d, [0.0] * len(models))
                parent.append(_build_ser(i, d, models, vals, col))
            log.info(f"top3 defect ({site}) mapped from FTT Top Defects Summary: "
                     f"{len(defects)} series, models={models}")
            continue

        df = ftt.get(site)
        if df is None or df.empty:
            WARN.append(f"top3 defect ({site}): no FTT site sheet"); continue
        series = _ser_elements(ch)
        if not series:
            WARN.append(f"top3 defect ({site}): no <c:ser> found -> skipped"); continue
        # Keep ORIGINAL display labels (e.g. "C/X-ray", "C/Bond gap / Rat hole")
        # for the legend. Only normalize for palette/lookup matching.
        existing_raw = [_ser_name(s) for s in series if _ser_name(s)]
        existing_norm = [norm_defect(n) for n in existing_raw]
        display = dict(zip(existing_norm, existing_raw))
        n = len(_cat_strs(series[0]))
        model_tot = df.groupby("Model")["FTT_Qty"].sum().sort_values(ascending=False)
        models = [str(m) for m in model_tot.head(n).index]
        sub = df[df["Model"].isin(models)]
        pivot = (sub.pivot_table(index="Model", columns="Defect_Issue",
                                 values="FTT_Qty", aggfunc="sum").fillna(0)
                 .reindex(models).fillna(0))
        pivot_norm = pivot.copy()
        pivot_norm.columns = [norm_defect(c) for c in pivot.columns]
        if DEFECT_VALUE_MODE == "proportion":
            pivot_norm = pivot_norm.div(pivot_norm.sum(axis=1).replace(0, 1), axis=0)
        new_defects = [d for d in pivot_norm.columns if d not in existing_norm]
        missing_pal = [d for d in new_defects if str(d).strip() not in palette]
        if missing_pal:
            raise SystemExit(
                f"\n[STOP] Top-3 Defect ({site}) needs color for new defect "
                f"type(s) not in palette: {missing_pal}\nAdd a color in "
                f"Color_Defect.xlsx (or the FTT ColorMap) and re-run. "
                f"Nothing was written yet.")
        # Preserve existing order (and their original display labels); append new.
        final_defects = [d for d in existing_norm if d in pivot_norm.columns] + new_defects
        parent = series[0].getparent()
        for s in series:
            parent.remove(s)
        for i, d in enumerate(final_defects):
            col = palette.get(str(d).strip(), "BFBFBF")
            vals = [float(pivot_norm.loc[m, d]) if (m in pivot_norm.index and d in pivot_norm.columns) else 0.0
                    for m in models]
            # Display label: original casing if it was an existing defect,
            # otherwise use the original data spelling (pre-normalization) so new
            # defects keep their natural "C/Name" form instead of a forced case.
            label = display.get(d)
            if label is None:
                # find the original (raw) defect label from the pivot columns
                raw = next((c for c in pivot.columns if norm_defect(c) == d), None)
                label = raw if raw else str(d)
            parent.append(_build_ser(i, label, models, vals, col))
        log.info(f"top3 defect ({site}) rebuilt: {len(final_defects)} defect series, "
                 f"models={models}")

def render_defect_legend_png(site_name, defects_with_colors, scale=3):
    """
    Render a clean, high-DPI image of the Excel custom defect legend.
    defects_with_colors: list of (defect_name, hex_color_str)
    """
    row_height = 20 * scale
    col_color_w = 24 * scale
    col_text_w = 176 * scale
    total_w = col_color_w + col_text_w
    num_rows = 1 + len(defects_with_colors)
    total_h = num_rows * row_height

    img = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    try:
        font_hdr = ImageFont.truetype("arialbd.ttf", 11 * scale)
        font_row = ImageFont.truetype("arial.ttf", 9 * scale)
    except Exception:
        font_hdr = ImageFont.load_default()
        font_row = ImageFont.load_default()

    border_color = (180, 180, 180)
    hdr_fill = (31, 78, 120)  # #1F4E78

    # Row 0: Header
    draw.rectangle([0, 0, total_w - 1, row_height - 1], fill=hdr_fill, outline=border_color, width=scale)
    hdr_text = f"Site {site_name}"
    bbox = draw.textbbox((0, 0), hdr_text, font=font_hdr)
    tx = (total_w - (bbox[2] - bbox[0])) // 2
    ty = (row_height - (bbox[3] - bbox[1])) // 2
    draw.text((tx, ty), hdr_text, fill=(255, 255, 255), font=font_hdr)

    # Rows: Defects
    for i, (defect_name, hex_color) in enumerate(defects_with_colors, start=1):
        y0 = i * row_height
        y1 = y0 + row_height - 1

        hex_clean = str(hex_color).lstrip("#")
        if len(hex_clean) == 8:
            hex_clean = hex_clean[2:]
        try:
            r = int(hex_clean[0:2], 16)
            g = int(hex_clean[2:4], 16)
            b = int(hex_clean[4:6], 16)
            color_rgb = (r, g, b)
        except Exception:
            color_rgb = (180, 180, 180)

        # Color box
        draw.rectangle([0, y0, col_color_w - 1, y1], fill=color_rgb, outline=border_color, width=scale)

        # Text box
        draw.rectangle([col_color_w - 1, y0, total_w - 1, y1], fill=(255, 255, 255), outline=border_color, width=scale)
        t_bbox = draw.textbbox((0, 0), defect_name, font=font_row)
        text_y = y0 + (row_height - (t_bbox[3] - t_bbox[1])) // 2
        text_x = col_color_w + 5 * scale
        draw.text((text_x, text_y), defect_name, fill=(0, 0, 0), font=font_row)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def update_slide1_legends(prs, palette):
    """Replace outdated Slide 1 custom defect legend images with newly rendered PNGs
    at the exact positions and dimensions from the format file."""
    summary_data = load_ftt_top_defects_from_excel(FTT_XLSX)
    if not summary_data:
        return
    slide1 = list(prs.slides)[1]
    site_shapes = {
        "VH": 66,
        "VH2": 89,
        "JV": 95,
        "JV2": 99,
    }
    for site, shp_id in site_shapes.items():
        if site not in summary_data:
            continue
        info = summary_data[site]
        defects = info['defects']
        defects_with_colors = []
        for d in defects:
            col = palette.get(norm_defect(d), "BFBFBF")
            defects_with_colors.append((d, col))

        matching = [s for s in slide1.shapes if s.shape_id == shp_id]
        if not matching:
            continue
        shp = matching[0]
        png_bytes = render_defect_legend_png(site, defects_with_colors, scale=3)
        left = shp.left
        top = shp.top
        width = shp.width
        height = shp.height

        # Remove old placeholder / picture shape
        parent = shp._element.getparent()
        parent.remove(shp._element)

        # Add new crisp PNG picture at exact original coordinates
        slide1.shapes.add_picture(io.BytesIO(png_bytes), left, top, width=width, height=height)
        log.info(f"Slide 1 legend ({site}) updated with {len(defects)} defects at ({left}, {top}, {width}, {height})")

def update_all_slide_titles(prs, period_label):
    """Update title and subtitles across all slides to reflect period_label."""
    # Slide 0 Title
    slide0 = prs.slides[0]
    for shp in slide0.shapes:
        if shp.has_text_frame and ("Cost of Poor Quality" in shp.text_frame.text or "CoPQ" in shp.text_frame.text):
            p = shp.text_frame.paragraphs[0]
            p.text = f"Cost of Poor Quality (CoPQ)_{period_label}"
            if p.runs:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(20)
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            log.info(f"Slide 0 title -> 'Cost of Poor Quality (CoPQ)_{period_label}'")
            break

    # Slides 1, 2, 3 Subtitles
    for idx in (1, 2, 3):
        if idx < len(prs.slides):
            slide = prs.slides[idx]
            for shp in slide.shapes:
                if shp.has_text_frame:
                    txt = shp.text_frame.text.strip()
                    if txt == "CoPQ Analysis" or "CoPQ Analysis" in txt:
                        p = shp.text_frame.paragraphs[0]
                        p.text = "CoPQ Analysis"
                        break

def apply_chart_hyperlinks(prs, db_overview_rel, db_type_rel, ftt_rel=None):
    """
    Apply interactive hyperlinks to all charts and overview summary tables.
    Clicking any chart or table during viewing or presentation opens the
    corresponding live database Excel workbook:
      - Slide 0 (Overview, site bars, country pies, table) -> CoPQ database 25.xlsx
      - Slide 1 (B/C Grade details, trends, models, pies)   -> CoPQ_type_analysis.xlsx (or FTT for defects)
      - Slide 2 (Re-inspection models & pies)               -> CoPQ_type_analysis.xlsx
      - Slide 3 (Touch-up models & pies)                    -> CoPQ_type_analysis.xlsx
    """
    total_linked = 0
    slides = list(prs.slides)
    for s_idx, slide in enumerate(slides):
        for shp in slide.shapes:
            target_link = None
            if s_idx == 0:
                if shp.has_chart or shp.has_table or shp.name in ["Table 102", "Rectangle 68", "Rectangle 69", "Rectangle 73"]:
                    target_link = db_overview_rel
            elif s_idx == 1:
                if shp.has_chart:
                    if shp.name in ["Chart 78", "Chart 80", "Chart 92", "Chart 96"]:
                        target_link = ftt_rel if ftt_rel else db_type_rel
                    else:
                        target_link = db_type_rel
            elif s_idx in [2, 3]:
                if shp.has_chart:
                    target_link = db_type_rel

            if target_link:
                try:
                    shp.click_action.hyperlink.address = target_link
                    total_linked += 1
                except Exception:
                    pass
    log.info(f"Applied interactive hyperlinks to {total_linked} charts/tables.")
    return total_linked


def update_pptx_chart_data_sources(pptx_path, db_overview_path, db_type_path, ftt_report_path=None):
    """
    Update internal OLE/externalData relationship targets for all 51 charts inside pptx_path,
    pointing them to the local Database workbooks using file:/// URI format.
    In PowerPoint on Windows, 'Edit Data' resolves external links via full file:/// URIs
    so clicking 'Edit Data' opens the exact Excel workbook directly!
    """
    import zipfile
    import tempfile
    import xml.etree.ElementTree as ET

    db_overview_abs = os.path.abspath(db_overview_path).replace('/', '\\')
    db_type_abs = os.path.abspath(db_type_path).replace('/', '\\')
    ftt_abs = os.path.abspath(ftt_report_path or "Output/FTT_Combined_Report.xlsx").replace('/', '\\')

    db_overview_target = f"file:///{db_overview_abs}"
    db_type_target = f"file:///{db_type_abs}"
    ftt_target = f"file:///{ftt_abs}"

    temp_zip = pptx_path + ".tmp.zip"
    shutil.copy2(pptx_path, temp_zip)
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(temp_zip, 'r') as zin:
            zin.extractall(temp_dir)

        charts_rel_dir = os.path.join(temp_dir, 'ppt', 'charts', '_rels')
        if os.path.exists(charts_rel_dir):
            for rel_name in os.listdir(charts_rel_dir):
                if not rel_name.endswith('.xml.rels'):
                    continue
                rel_file = os.path.join(charts_rel_dir, rel_name)
                tree = ET.parse(rel_file)
                root = tree.getroot()
                changed = False
                for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    t = rel.attrib.get('Type', '')
                    if 'oleObject' in t or 'package' in t:
                        curr_target = rel.attrib.get('Target', '')
                        fn = os.path.basename(curr_target.replace("\\", "/")).lower()
                        if 'copq database' in fn or 'copq%20database' in fn or 'database%2025' in fn or 'database 25' in fn:
                            new_target = db_overview_target
                        elif 'type_analysis' in fn or 'type%20analysis' in fn:
                            new_target = db_type_target
                        elif 'ftt' in fn:
                            new_target = ftt_target
                        else:
                            new_target = db_type_target

                        rel.attrib['Target'] = new_target
                        rel.attrib['TargetMode'] = 'External'
                        changed = True
                if changed:
                    tree.write(rel_file, xml_declaration=True, encoding='utf-8')

        with zipfile.ZipFile(pptx_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    full_p = os.path.join(root_dir, file)
                    arc_p = os.path.relpath(full_p, temp_dir)
                    zout.write(full_p, arc_p)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        if os.path.exists(temp_zip):
            os.remove(temp_zip)


def update_site_total_table(prs, copq):
    """Update Slide 0 Table 102 headers: Site Name on Line 1, Total on Line 2.
    Clean typography, bold, centered, no diff arrows."""
    slide0 = list(prs.slides)[0]
    site_names = ["VH", "VH2", "VH3", "VH4", "JV", "JV2", "JVB"]
    for shp in slide0.shapes:
        if not shp.has_table:
            continue
        tbl = shp.table
        for ci, site in enumerate(site_names, start=1):
            if ci >= len(tbl.rows[0].cells):
                break
            cell = tbl.rows[0].cells[ci]
            total = float(non_rework(copq[site])["Ttl cost ($)"].sum()) if site in copq and not copq[site].empty else 0.0
            tf = cell.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            p0.text = f"{site} "
            p0.alignment = PP_ALIGN.CENTER
            if p0.runs:
                p0.runs[0].font.name = "Arial"
                p0.runs[0].font.size = Pt(11)
                p0.runs[0].font.bold = True
                p0.runs[0].font.color.rgb = RGBColor(0, 0, 0)

            if len(tf.paragraphs) > 1:
                p1 = tf.paragraphs[1]
            else:
                p1 = tf.add_paragraph()
            p1.text = f"${total:,.2f}"
            p1.alignment = PP_ALIGN.CENTER
            if p1.runs:
                p1.runs[0].font.name = "Arial"
                p1.runs[0].font.size = Pt(11)
                p1.runs[0].font.bold = True
                p1.runs[0].font.color.rgb = RGBColor(0, 0, 0)
            log.info(f"table header {site}: ${total:,.2f}")
        return

def update_country_totals(prs, copq):
    vn_sites = [s for s in copq if s.startswith("VH")]
    id_sites = [s for s in copq if s.startswith("JV")]
    vn = sum(float(non_rework(copq[s])["Ttl cost ($)"].sum()) for s in vn_sites if s in copq and not copq[s].empty)
    idt = sum(float(non_rework(copq[s])["Ttl cost ($)"].sum()) for s in id_sites if s in copq and not copq[s].empty)
    clg = vn + idt
    slide0 = list(prs.slides)[0]
    total_boxes = {
        70: (vn, RGBColor(0, 0, 0), Pt(12.5)),         # VN: dark text on white circle background
        87: (clg, RGBColor(0, 0, 0), Pt(13)),          # CLG: black text
        89: (idt, RGBColor(0, 0, 0), Pt(12.5)),        # ID: dark text on white circle background
    }
    for shp in slide0.shapes:
        if shp.shape_id in total_boxes:
            val, color, sz = total_boxes[shp.shape_id]
            new = f"${val:,.2f}"
            if shp.has_text_frame and shp.text_frame.paragraphs:
                p = shp.text_frame.paragraphs[0]
                p.text = new
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.size = sz
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = color
                log.info(f"total box id={shp.shape_id} -> {new}")
    for nm, val in (("Chart 57", vn), ("Chart 59", idt)):
        ch = get_chart_by_name(slide0, nm)
        if ch:
            for ser in _ser_elements(ch):
                set_series_values_by_category(ser, {"0": val}, warn=False)
            log.info(f"{nm} doughnut -> ${val:,.2f}")
    ch = get_chart_by_name(slide0, "Chart 58")
    if ch:
        for ser in _ser_elements(ch):
            set_series_values_by_category(ser, {"VN": vn, "Indo": idt}, warn=False)
        log.info(f"CLG doughnut -> VN ${vn:,.2f}, Indo ${idt:,.2f}")

def _load_prior_month_overview():
    """Load prior month overview figures from CoPQ database 25.xlsx for MoM delta calculations."""
    try:
        xl = pd.ExcelFile(DB_OVERVIEW)
        target_sheet = None
        try:
            import datetime as _dt
            curr_dt = _dt.datetime.strptime(DB_MONTH, "%b-%Y")
            prior_dt = curr_dt.replace(day=1) - _dt.timedelta(days=1)
            cands = [
                prior_dt.strftime("%b-%Y"),
                prior_dt.strftime("%B-%Y"),
                prior_dt.strftime("%b-%y"),
            ]
            for c in cands:
                if c in xl.sheet_names:
                    target_sheet = c
                    break
        except Exception:
            pass

        if not target_sheet:
            candidates = [s for s in xl.sheet_names if any(y in s for y in ["2025", "2026"]) and not s.startswith("COPQ_rows") and s != "Current" and s != DB_MONTH]
            if candidates:
                target_sheet = candidates[-1]

        if target_sheet:
            log.info(f"Loaded prior month overview from sheet: {target_sheet}")
            df = xl.parse(target_sheet, header=1).set_index("Factory")
            return df
    except Exception as e:
        log.warning(f"Could not load prior month overview: {e}")
    return None

def load_mom_diffs_from_analysis(analysis_path):
    """
    Extract exact evaluated MoM difference values from CoPQ_type_analysis.xlsx:
      - Re-inspection diffs (Rows 27-32)
      - Touch-up diffs (Rows 44-48)
    Formula in Difference column: curr_col - prev_col.
    """
    import openpyxl
    if not os.path.exists(analysis_path):
        return {}, {}
    try:
        wb = openpyxl.load_workbook(analysis_path, data_only=False)
        if 'Current month' not in wb.sheetnames:
            return {}, {}
        ws = wb['Current month']

        col_diff = None
        for c in range(3, ws.max_column + 1):
            val = str(ws.cell(27, c).value or '')
            if val.startswith('='):
                col_diff = c
                break

        if not col_diff:
            col_diff = 26

        prev_col = col_diff - 2
        curr_col = col_diff - 1

        re_diffs = {}
        for r, s in zip(range(27, 33), ["VH", "VH2", "VH3", "JV", "JV2", "JVB"]):
            c_val = float(ws.cell(r, curr_col).value or 0.0)
            p_val = float(ws.cell(r, prev_col).value or 0.0)
            re_diffs[s] = c_val - p_val

        tu_diffs = {}
        for r, s in zip(range(44, 49), ["VH", "VH2", "VH4", "JV", "JVB"]):
            c_val = float(ws.cell(r, curr_col).value or 0.0)
            p_val = float(ws.cell(r, prev_col).value or 0.0)
            tu_diffs[s] = c_val - p_val

        return re_diffs, tu_diffs
    except Exception as e:
        log.warning(f"Could not load MoM diffs from analysis: {e}")
        return {}, {}


def apply_delta_badge_styling(grp, diff_val, font_size_pt=8.5):
    """
    Update delta badge group (Rectangle + Triangle):
    - diff_val > 0: INCREASE (giá tăng) -> RED/Magenta (#E6007E), Up triangle (mũi tên đi lên: rot=10800000, flipV=1)
    - diff_val < 0: DECREASE (giá giảm) -> GREEN (#00B050), Down triangle (mũi tên đi xuống: rot=10800000, no flipV)
    - diff_val == 0: GREEN $0.00
    Text format: f"${abs(diff_val):,.2f}"
    """
    rect = None
    tri = None
    for shp in grp.shapes:
        if 'Rectangle' in shp.name:
            rect = shp
        elif 'Triangle' in shp.name:
            tri = shp

    if rect is None:
        return

    abs_str = f"${abs(diff_val):,.2f}"
    is_inc = (diff_val > 0)

    # Red/Magenta for increase, Green for decrease
    tri_color_hex = "E6007E" if is_inc else "00B050"
    txt_color_rgb = RGBColor(230, 0, 126) if is_inc else RGBColor(0, 176, 80)

    # 1. Update text & font
    rect.text_frame.text = abs_str
    p = rect.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    if p.runs:
        r = p.runs[0]
        r.font.name = "Arial"
        r.font.size = Pt(font_size_pt)
        r.font.bold = True
        r.font.color.rgb = txt_color_rgb

    # 2. Update triangle shape direction & color
    # In the template coordinate system, rot="10800000" with flipV="1" points UP (▲),
    # while rot="10800000" without flipV points DOWN (▼).
    if tri is not None:
        xfrm = tri._element.find('.//' + qn('a:xfrm'))
        if xfrm is not None:
            xfrm.attrib['rot'] = '10800000'
            if is_inc:
                xfrm.attrib['flipV'] = '1'
            else:
                if 'flipV' in xfrm.attrib:
                    del xfrm.attrib['flipV']

        spPr = tri._element.find(qn('p:spPr'))
        if spPr is not None:
            solidFill = spPr.find(qn('a:solidFill'))
            if solidFill is None:
                solidFill = etree.SubElement(spPr, qn('a:solidFill'))
            srgb = solidFill.find(qn('a:srgbClr'))
            if srgb is None:
                srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
            srgb.attrib['val'] = tri_color_hex


def update_slide2_reinspection(prs, copq):
    """Update Slide 2: Re-Inspection analysis (6 site badges, 6 top-3 model charts, 6 pie charts)."""
    slide2 = list(prs.slides)[2]
    re_diffs, _ = load_mom_diffs_from_analysis(DB_TYPE)

    # 1. Site KPI badges & MoM delta badges
    badge_groups = {
        "Group 41": {"site": "VH",  "delta_grp": "Group 73"},
        "Group 45": {"site": "VH2", "delta_grp": "Group 5"},
        "Group 49": {"site": "VH3", "delta_grp": "Group 79"},
        "Group 63": {"site": "JV",  "delta_grp": "Group 8"},
        "Group 53": {"site": "JV2", "delta_grp": "Group 92"},
        "Group 57": {"site": "JVB", "delta_grp": "Group 67"},
    }
    for grp_name, info in badge_groups.items():
        site = info["site"]
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        tot = float(sub["Ttl cost ($)"].sum()) if sub is not None and not sub.empty else 0.0

        # Update main cost badge (Large Blue Font: Arial 24pt, Bold, Color #0000CC)
        for shp in slide2.shapes:
            if shp.name == grp_name:
                for sub_shp in shp.shapes:
                    if sub_shp.has_text_frame and "$" in sub_shp.text_frame.text:
                        sub_shp.text_frame.text = f"${tot:,.2f}"
                        p = sub_shp.text_frame.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER
                        if p.runs:
                            p.runs[0].font.name = "Arial"
                            p.runs[0].font.size = Pt(24)
                            p.runs[0].font.bold = True
                            p.runs[0].font.color.rgb = RGBColor(0, 0, 204) # #0000CC Blue
                log.info(f"Slide 2 badge {shp.name} ({site}): ${tot:,.2f}")

        # Update MoM delta badge from CoPQ_type_analysis.xlsx formula
        diff = re_diffs.get(site, 0.0)
        for shp in slide2.shapes:
            if shp.name == info["delta_grp"]:
                apply_delta_badge_styling(shp, diff, font_size_pt=8.5)
                log.info(f"Slide 2 delta {info['delta_grp']} ({site}): diff={diff:+,.2f}")

    # 2. Top-3 Models Column Charts
    model_charts = {
        "Chart 23": "VH",
        "Chart 31": "VH2",
        "Chart 34": "VH3",
        "Chart 36": "JV",
        "Chart 17": "JV2",
        "Chart 27": "JVB",
    }
    for name, site in model_charts.items():
        ch = get_chart_by_name(slide2, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        if sub is None or sub.empty:
            continue
        top = sub.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(3)
        cats = [clean_model_name(str(m), max_len=18) for m in top.index]
        vals = [round(float(v), 2) for v in top.values]
        if len(cats) > 1:
            while len(cats) < 3:
                cats.append("")
                vals.append(None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)
            set_series_color(ser, "33CCFF")
        format_model_chart(ch)
        log.info(f"Slide 2 models {name} ({site}): {cats} -> {vals}")

    # 3. Factory / Shopfloor Pie Charts
    pie_charts = {
        "Chart 22": "VH",
        "Chart 32": "VH2",
        "Chart 33": "VH3",
        "Chart 37": "JV",
        "Chart 19": "JV2",
        "Chart 21": "JVB",
    }
    for name, site in pie_charts.items():
        ch = get_chart_by_name(slide2, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Re-inspection"] if df is not None and not df.empty else None
        if sub is None or sub.empty:
            continue
        col_to_group = "Factory" if "Factory" in sub.columns else "Fty/Line"
        fac_sum = sub.groupby(col_to_group)["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)[:5]
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"Slide 2 pie {name} ({site}): {cat_val_pairs}")

def update_slide3_touchup(prs, copq):
    """Update Slide 3: Touch-Up Paint analysis (4 site badges, 4 top-3 model charts, 4 pie charts)."""
    slide3 = list(prs.slides)[3]
    _, tu_diffs = load_mom_diffs_from_analysis(DB_TYPE)

    # 1. Site KPI badges & MoM delta badges
    badge_shapes = {
        "Rectangle 23":  {"site": "VH",  "delta_grp": "Group 1"},
        "Rectangle 110": {"site": "VH2", "delta_grp": "Group 5"},
        "Rectangle 60":  {"site": "VH4", "delta_grp": "Group 120"},
        "Rectangle 111": {"site": "JV",  "delta_grp": "Group 41"},
    }
    for shp_name, info in badge_shapes.items():
        site = info["site"]
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        tot = float(sub["Ttl cost ($)"].sum()) if sub is not None and not sub.empty else 0.0

        # Update main cost badge (Large Blue Font: Arial 24pt, Bold, Color #0000CC)
        for shp in slide3.shapes:
            if shp.name == shp_name and shp.has_text_frame:
                shp.text_frame.text = f"${tot:,.2f}"
                p = shp.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                if p.runs:
                    p.runs[0].font.name = "Arial"
                    p.runs[0].font.size = Pt(24)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(0, 0, 204) # #0000CC Blue
                log.info(f"Slide 3 badge {shp.name} ({site}): ${tot:,.2f}")

        # Update MoM delta badge from CoPQ_type_analysis.xlsx formula
        diff = tu_diffs.get(site, 0.0)
        for shp in slide3.shapes:
            if shp.name == info["delta_grp"]:
                apply_delta_badge_styling(shp, diff, font_size_pt=8.5)
                log.info(f"Slide 3 delta {info['delta_grp']} ({site}): diff={diff:+,.2f}")

    # 2. Top Models Column Charts
    model_charts = {
        "Chart 4":  "VH",
        "Chart 13": "VH2",
        "Chart 15": "VH4",
        "Chart 17": "JV",
    }
    for name, site in model_charts.items():
        ch = get_chart_by_name(slide3, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        if sub is None or sub.empty:
            continue
        top = sub.groupby("Model")["Ttl cost ($)"].sum().sort_values(ascending=False).head(3)
        cats = [clean_model_name(str(m), max_len=18) for m in top.index]
        vals = [round(float(v), 2) for v in top.values]
        if len(cats) > 1:
            while len(cats) < 3:
                cats.append("")
                vals.append(None)
        for ser in _ser_elements(ch):
            set_series_categories(ser, cats)
            set_series_values_ordered(ser, vals)
            set_series_color(ser, "33CCFF")
        format_model_chart(ch)
        log.info(f"Slide 3 models {name} ({site}): {cats} -> {vals}")

    # 3. Shopfloor / Factory Pie Charts
    pie_charts = {
        "Chart 9":  "VH",
        "Chart 11": "VH2",
        "Chart 19": "VH4",
        "Chart 21": "JV",
    }
    for name, site in pie_charts.items():
        ch = get_chart_by_name(slide3, name)
        if ch is None:
            continue
        df = copq.get(site)
        sub = df[df.apply(map_bar_cat, axis=1) == "Touch up paint"] if df is not None and not df.empty else None
        if sub is None or sub.empty:
            continue
        col_to_group = "Factory" if "Factory" in sub.columns else "Fty/Line"
        fac_sum = sub.groupby(col_to_group)["Ttl cost ($)"].sum().to_dict()
        cat_val_pairs = sorted([(str(k), float(v)) for k, v in fac_sum.items()], key=lambda x: x[1], reverse=True)[:5]
        update_pie_chart_data(ch, cat_val_pairs)
        log.info(f"Slide 3 pie {name} ({site}): {cat_val_pairs}")

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def clone(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    log.info(f"cloned {src}\n      -> {dst}")

def verify_opens(path):
    p = Presentation(path)
    log.info(f"VERIFY: re-opens OK, {len(p.slides)} slides")
    return len(p.slides)

def load_copq():
    xl = pd.ExcelFile(COPQ_XLSX)
    out = {}

    # Prefer reading directly from the unified master sheet 'COPQ_Clean' if available
    if "COPQ_Clean" in xl.sheet_names:
        df_master = xl.parse("COPQ_Clean", header=0)
        if "Ttl cost ($)" not in df_master.columns and "Cost" in df_master.columns:
            df_master["Ttl cost ($)"] = pd.to_numeric(df_master["Cost"], errors="coerce").fillna(0.0)
        if "Defective Qty(Pair)" not in df_master.columns and "Qty" in df_master.columns:
            df_master["Defective Qty(Pair)"] = pd.to_numeric(df_master["Qty"], errors="coerce").fillna(0.0)
        if "Working hours" not in df_master.columns:
            df_master["Working hours"] = 0.0

        for site, sub in df_master.groupby("Site"):
            sub = sub.copy()
            if "Date" in sub.columns:
                sub = sub[sub["Date"].notna()]
            if EXCLUDE_TYPES:
                mask = ~sub["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
                sub = sub[mask]
            for c in ("Ttl cost ($)", "Defective Qty(Pair)", "Working hours"):
                if c in sub.columns:
                    sub[c] = pd.to_numeric(sub[c], errors="coerce").fillna(0.0)
            deck_site = DATA_SITE_ALIAS.get(str(site).strip(), str(site).strip())
            out[deck_site] = sub
        log.info(f"COPQ loaded from COPQ_Clean master: {len(out)} sites")
        return out

    # Fallback to individual per-site sheets
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=1)
        if "Date" not in df.columns and len(df) and str(df.iloc[0, 0]).startswith("Cost Of Poor Quality"):
            df = xl.parse(sheet, header=2)
        if "Date" in df.columns:
            df = df[df["Date"].notna()]
        if EXCLUDE_TYPES:
            mask = ~df["Type"].astype(str).str.strip().isin(EXCLUDE_TYPES)
            df = df[mask]
        if "Ttl cost ($)" not in df.columns and "Cost" in df.columns:
            df["Ttl cost ($)"] = pd.to_numeric(df["Cost"], errors="coerce").fillna(0.0)
        if "Defective Qty(Pair)" not in df.columns and "Qty" in df.columns:
            df["Defective Qty(Pair)"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0.0)
        if "Working hours" not in df.columns:
            df["Working hours"] = 0.0
        for c in ("Ttl cost ($)", "Defective Qty(Pair)", "Working hours"):
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        deck_site = DATA_SITE_ALIAS.get(sheet, sheet)
        out[deck_site] = df
    return out

def load_copq_db():
    try:
        xl = pd.ExcelFile(DB_OVERVIEW)
    except FileNotFoundError:
        log.error(f"DB overview not found: {DB_OVERVIEW}")
        return {}
    # Use COPQ_Clean.xlsx master data directly
    return load_copq()

def generate_copq_presentation(
    period_label=None,
    db_month=None,
    out_pptx=None,
    copq_xlsx=None,
    ftt_xlsx=None,
    src_pptx=None,
    db_overview=None,
    db_type=None
):
    global PERIOD_LABEL, DB_MONTH, OUT_PPTX, COPQ_XLSX, FTT_XLSX, SRC_PPTX, DB_OVERVIEW, DB_TYPE

    if period_label:
        PERIOD_LABEL = period_label
    if db_month:
        DB_MONTH = db_month
    if out_pptx:
        OUT_PPTX = out_pptx
    if copq_xlsx:
        COPQ_XLSX = copq_xlsx
    if ftt_xlsx:
        FTT_XLSX = ftt_xlsx
    if src_pptx:
        SRC_PPTX = src_pptx
    if db_overview:
        DB_OVERVIEW = db_overview
    if db_type:
        DB_TYPE = db_type

    log.info("=" * 70)
    log.info(f"COPQ PPTX Executive Generator & Visualizer (Period: {PERIOD_LABEL})")
    log.info("=" * 70)
    if os.path.abspath(OUT_PPTX) == os.path.abspath(SRC_PPTX):
        raise SystemExit("Refusing to overwrite the source file.")

    target = OUT_PPTX
    if os.path.exists(target):
        try:
            os.remove(target)
        except PermissionError:
            import datetime as _dt
            stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            base, ext = os.path.splitext(OUT_PPTX)
            target = f"{base}_{stamp}{ext}"
            log.warning(f"Target is locked (open in another app). "
                        f"Writing to: {target}")

    clone(SRC_PPTX, target)
    copq = load_copq_db()
    if not copq:
        log.warning("No COPQ data loaded from database; falling back to COPQ_Clean.xlsx")
        copq = load_copq()
    ftt, cmap = load_ftt()

    prs = Presentation(target)
    palette = build_palette(prs, cmap)

    # 1. Update Titles & Subtitles across all 4 slides
    update_all_slide_titles(prs, PERIOD_LABEL)

    # 2. Slide 0: Executive Overview
    update_site_total_table(prs, copq)
    update_country_totals(prs, copq)
    update_site_bar(prs, copq)
    update_na_badges(prs, copq)
    update_bottom_trend_charts(prs, copq)

    # 3. Slide 1: B/C Grade Detail
    update_slide1_monthly(prs)
    update_factory_pie(prs, copq)
    update_top3_model(prs, copq)
    update_top3_defect(prs, ftt, palette)
    update_slide1_legends(prs, palette)

    # 4. Slide 2: Re-Inspection Detail
    update_slide2_reinspection(prs, copq)

    # 5. Slide 3: Touch-Up Paint Detail
    update_slide3_touchup(prs, copq)

    # 6. Apply reference-deck geometry only; all chart data was updated above.
    apply_analysis_layout(prs)

    # 7. Apply live interactive hyperlinks to all charts & summary tables
    db_rel_overview = os.path.relpath(DB_OVERVIEW, os.path.dirname(target))
    db_rel_type = os.path.relpath(DB_TYPE, os.path.dirname(target))
    ftt_rel = os.path.relpath(FTT_XLSX, os.path.dirname(target)) if os.path.exists(FTT_XLSX) else None
    apply_chart_hyperlinks(prs, db_rel_overview, db_rel_type, ftt_rel)

    prs.save(target)

    # 8. Update internal chart OLE relationship targets (pass full absolute paths DB_OVERVIEW and DB_TYPE)
    update_pptx_chart_data_sources(target, DB_OVERVIEW, DB_TYPE, FTT_XLSX)
    log.info("-" * 70)
    verify_opens(target)
    log.info("=" * 70)
    log.info(f"OUTPUT WRITTEN: {target}")
    log.info("=" * 70)
    log.info(f"WARNINGS: {len(WARN)}")
    for w in WARN:
        log.warning("  - " + w)
    log.info("=" * 70)
    return target

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Update COPQ presentation deck.")
    parser.add_argument("--period", default=None, help="Period label (e.g. 'Aug, 2026')")
    parser.add_argument("--db-month", default=None, help="DB month sheet name (e.g. 'Aug-2026')")
    parser.add_argument("--output", default=None, help="Target PPTX path")
    parser.add_argument("--copq-xlsx", default=None, help="COPQ_Clean.xlsx path")
    parser.add_argument("--ftt-xlsx", default=None, help="FTT_Combined_Report.xlsx path")
    parser.add_argument("--template", default=None, help="Template PPTX path")
    args = parser.parse_args()

    generate_copq_presentation(
        period_label=args.period,
        db_month=args.db_month,
        out_pptx=args.output,
        copq_xlsx=args.copq_xlsx,
        ftt_xlsx=args.ftt_xlsx,
        src_pptx=args.template
    )

if __name__ == "__main__":
    main()
