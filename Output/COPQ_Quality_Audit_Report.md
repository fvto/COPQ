# COPQ Quality Health & Anomaly Audit Report
**Audit Timestamp:** 2026-10-02T10:57:03.366155  
**Reporting Period:** Sep, 2026  
**Overall Health Score:** `90/100`  
**System Status:** `HEALTHY`  

---

## 1. Executive Summary & Action Items

| Severity | Category | Operational Finding / Recommendation |
|---|---|---|
| 🟡 **WARNING** | Defect Dictionary | Found 27 defect(s) in active FTT logs not present in Color_Defect.xlsx: B/Holes quality, B/Jump / Broken / Loose Stitch, B/Material damaged, B/Off-center, B/Other defects.... These may trigger fallback colors in PPTX charts. |

## 2. Multi-Plant Site Coverage

- **COPQ ERP Exports:** 7 sites detected (`JV, JV2, JV3, VH, VH2, VH3, VH4`).
- **Missing COPQ Sites:** None (100% Complete).
- **FTT B/C Defect Reports:** 4 sites detected (`JV, JV2, VH, VH2`).

## 3. Defect Dictionary & Color Mapping Integrity

- **Palette Defect Terms in SOP:** 15
- **Active Defects Extracted:** 27
- **Unmapped Defects:** 27

**Unmapped Defect Terms:**
- `B/Holes quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Jump / Broken / Loose Stitch` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Material damaged` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Off-center` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Other defects` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Stitching margin` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Wrinkle or Deformed Upper` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/X-ray` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Airbag defect` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Alignment L+R Symmetry` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Bond gap / Rat hole` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Color / Paint migration, bleeding` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Color mis-match` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Contamination` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Delamination` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Holes quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Interior defect` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Jump / Broken / Loose Stitch` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Material damaged` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Off-center` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Other defects` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Over buffing` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Paint surface quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Stitching margin` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Wrinkle or Deformed Upper` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/X-ray` -> *Assigned fallback palette color in openpyxl/pptx*
- `C/Yellowing` -> *Assigned fallback palette color in openpyxl/pptx*

## 4. Statistical & Heuristic Anomaly Audit

- **Total Master Records Audited:** 2,259
- **High Working Hours (>15h):** 0 rows
- **High Single-Event Cost (>$1,000):** 0 rows
- **Negative Quantities / Cost:** 0 rows

---
*Report generated proactively by QualitySentinelAgent.*