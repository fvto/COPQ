# COPQ Quality Health & Anomaly Audit Report
**Audit Timestamp:** 2026-10-02T11:18:38.815539  
**Reporting Period:** Current  
**Overall Health Score:** `80/100`  
**System Status:** `WARNINGS_FOUND`  

---

## 1. Executive Summary & Action Items

| Severity | Category | Operational Finding / Recommendation |
|---|---|---|
| 🟡 **WARNING** | Defect Dictionary | Found 128 defect(s) in active FTT logs not present in Color_Defect.xlsx: Accessories attachment, Aglet deform, Airbag quality/stain, B/Holes quality, B/Jump / Broken / Loose Stitch.... These may trigger fallback colors in PPTX charts. |
| 🟡 **WARNING** | Working Hours Anomaly | 1 records exceed standard working hours limit (> 15.0h). Samples: AIR JORDAN 1 MID GS. |

## 2. Multi-Plant Site Coverage

- **COPQ ERP Exports:** 7 sites detected (`JV, JV2, JV3, VH, VH2, VH3, VH4`).
- **Missing COPQ Sites:** None (100% Complete).
- **FTT B/C Defect Reports:** 4 sites detected (`JV, JV2, VH, VH2`).

## 3. Defect Dictionary & Color Mapping Integrity

- **Palette Defect Terms in SOP:** 15
- **Active Defects Extracted:** 128
- **Unmapped Defects:** 128

**Unmapped Defect Terms:**
- `Accessories attachment` -> *Assigned fallback palette color in openpyxl/pptx*
- `Aglet deform` -> *Assigned fallback palette color in openpyxl/pptx*
- `Airbag quality/stain` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Holes quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Jump / Broken / Loose Stitch` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Material damaged` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Off-center` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Other defects` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Stitching margin` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/Wrinkle or Deformed Upper` -> *Assigned fallback palette color in openpyxl/pptx*
- `B/X-ray` -> *Assigned fallback palette color in openpyxl/pptx*
- `Binding` -> *Assigned fallback palette color in openpyxl/pptx*
- `Blanket stitch` -> *Assigned fallback palette color in openpyxl/pptx*
- `Blooming` -> *Assigned fallback palette color in openpyxl/pptx*
- `Bond gap at bottom` -> *Assigned fallback palette color in openpyxl/pptx*
- `Bottom Cosmetic` -> *Assigned fallback palette color in openpyxl/pptx*
- `Bottom long/short` -> *Assigned fallback palette color in openpyxl/pptx*
- `Box stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Broken thread/Run off stitching` -> *Assigned fallback palette color in openpyxl/pptx*
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
- `Cleanness` -> *Assigned fallback palette color in openpyxl/pptx*
- `Close seam line unstraight` -> *Assigned fallback palette color in openpyxl/pptx*
- `Collapse vamp` -> *Assigned fallback palette color in openpyxl/pptx*
- `Collar pillow/bootie shape` -> *Assigned fallback palette color in openpyxl/pptx*
- `Collar shape` -> *Assigned fallback palette color in openpyxl/pptx*
- `Color bleeding` -> *Assigned fallback palette color in openpyxl/pptx*
- `Color matching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Color migration` -> *Assigned fallback palette color in openpyxl/pptx*
- `Components inconsistent` -> *Assigned fallback palette color in openpyxl/pptx*
- `Computer stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Cutting holes quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `Damaged lace/wrong lacing` -> *Assigned fallback palette color in openpyxl/pptx*
- `Damaged material/delamination` -> *Assigned fallback palette color in openpyxl/pptx*
- `Damaged/scratch` -> *Assigned fallback palette color in openpyxl/pptx*
- `Deco-stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Definition after process` -> *Assigned fallback palette color in openpyxl/pptx*
- `Deformed bottom` -> *Assigned fallback palette color in openpyxl/pptx*
- `Distance deco-stitch to biteline` -> *Assigned fallback palette color in openpyxl/pptx*
- `Distance logo to biteline` -> *Assigned fallback palette color in openpyxl/pptx*
- `Double stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Elastic band stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Embroidery stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Exposed reinforcement` -> *Assigned fallback palette color in openpyxl/pptx*
- `Eyelet` -> *Assigned fallback palette color in openpyxl/pptx*
- `Eyestay off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Eyestay/Collar Opening` -> *Assigned fallback palette color in openpyxl/pptx*
- `Fold U-shape eyestay` -> *Assigned fallback palette color in openpyxl/pptx*
- `Folding` -> *Assigned fallback palette color in openpyxl/pptx*
- `HFW bond gap` -> *Assigned fallback palette color in openpyxl/pptx*
- `Hairy edge` -> *Assigned fallback palette color in openpyxl/pptx*
- `Heel off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Heel shape` -> *Assigned fallback palette color in openpyxl/pptx*
- `Heel stitching line up` -> *Assigned fallback palette color in openpyxl/pptx*
- `Heel/Tip height inconsistent` -> *Assigned fallback palette color in openpyxl/pptx*
- `Hot-melt over flow` -> *Assigned fallback palette color in openpyxl/pptx*
- `Jump man logo quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `Lace cleanness` -> *Assigned fallback palette color in openpyxl/pptx*
- `Lacing` -> *Assigned fallback palette color in openpyxl/pptx*
- `Line up upper&bottom` -> *Assigned fallback palette color in openpyxl/pptx*
- `Loose stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Material color shade` -> *Assigned fallback palette color in openpyxl/pptx*
- `Midsole/Outsole to upper bond gap` -> *Assigned fallback palette color in openpyxl/pptx*
- `Nosew bond gap` -> *Assigned fallback palette color in openpyxl/pptx*
- `Nosew bubble` -> *Assigned fallback palette color in openpyxl/pptx*
- `Nosew over cement` -> *Assigned fallback palette color in openpyxl/pptx*
- `Others` -> *Assigned fallback palette color in openpyxl/pptx*
- `Outsole thread stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Over buffing` -> *Assigned fallback palette color in openpyxl/pptx*
- `Over cement` -> *Assigned fallback palette color in openpyxl/pptx*
- `Over painting` -> *Assigned fallback palette color in openpyxl/pptx*
- `Paint peeled off` -> *Assigned fallback palette color in openpyxl/pptx*
- `Punching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Quality of cork` -> *Assigned fallback palette color in openpyxl/pptx*
- `Rocking` -> *Assigned fallback palette color in openpyxl/pptx*
- `Serge stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Size label peeled off` -> *Assigned fallback palette color in openpyxl/pptx*
- `Sockliner position` -> *Assigned fallback palette color in openpyxl/pptx*
- `Sockliner quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `Sole attachment` -> *Assigned fallback palette color in openpyxl/pptx*
- `Stitching margin/SPI` -> *Assigned fallback palette color in openpyxl/pptx*
- `Stitching wrong component` -> *Assigned fallback palette color in openpyxl/pptx*
- `Strap collapses` -> *Assigned fallback palette color in openpyxl/pptx*
- `Strap height inconsistent` -> *Assigned fallback palette color in openpyxl/pptx*
- `Strap stitching` -> *Assigned fallback palette color in openpyxl/pptx*
- `Swoosh quality` -> *Assigned fallback palette color in openpyxl/pptx*
- `The label is upside down` -> *Assigned fallback palette color in openpyxl/pptx*
- `Thread end` -> *Assigned fallback palette color in openpyxl/pptx*
- `Toe cap shape` -> *Assigned fallback palette color in openpyxl/pptx*
- `Toe collapse` -> *Assigned fallback palette color in openpyxl/pptx*
- `Toe off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Toe spring` -> *Assigned fallback palette color in openpyxl/pptx*
- `Tongue folding` -> *Assigned fallback palette color in openpyxl/pptx*
- `Tongue off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Tongue shape` -> *Assigned fallback palette color in openpyxl/pptx*
- `Trace on upper` -> *Assigned fallback palette color in openpyxl/pptx*
- `Trimming` -> *Assigned fallback palette color in openpyxl/pptx*
- `Webbing off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Woven label off center` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle Counter & lump` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle bottom` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle collar` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle collar lining` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle heel` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle lining` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle toe cap` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle tongue` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrinkle upper` -> *Assigned fallback palette color in openpyxl/pptx*
- `Wrong style` -> *Assigned fallback palette color in openpyxl/pptx*
- `X-ray` -> *Assigned fallback palette color in openpyxl/pptx*
- `Yellowing` -> *Assigned fallback palette color in openpyxl/pptx*

## 4. Statistical & Heuristic Anomaly Audit

- **Total Master Records Audited:** 2,259
- **High Working Hours (>15h):** 1 rows
- **High Single-Event Cost (>$1,000):** 0 rows
- **Negative Quantities / Cost:** 0 rows

---
*Report generated proactively by QualitySentinelAgent.*