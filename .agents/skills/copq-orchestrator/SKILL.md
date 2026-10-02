---
name: copq-orchestrator
description: >-
  Autonomous orchestrator, data quality sentinel, and executive delivery suite
  for the COPQ (Cost of Poor Quality) and FTT (First-Time-Through) manufacturing
  analytics system. Use when processing ERP exports, auditing data quality,
  checking defect color mapping, or compiling executive presentations and briefings.
---

# COPQ Autonomous Pipeline & Sentinel Skill

This skill provides operational runbooks and guidelines for running the 3 proactive agents:
1. **Pipeline Orchestrator Agent** (`agents/pipeline_orchestrator.py`)
2. **Quality Sentinel & Anomaly Triage Agent** (`agents/quality_sentinel.py`)
3. **Watchdog & Executive Delivery Agent** (`agents/watchdog_delivery.py`)

---

## 🚀 Quick Execution Commands

From the workspace root (`combine_COPQ`):

### 1. Run Complete Autonomous Cycle (Audit -> Orchestrate -> Brief)
```bash
python agents/run_agent.py run-all --period "2026-08"
```
Or with auto-detected period:
```bash
python agents/run_agent.py run-all
```

### 2. Run Pre-flight Data Quality & Anomaly Audit
```bash
python agents/run_agent.py audit
```
Inspect generated outputs:
- `Output/COPQ_Quality_Audit_Report.md`
- `Output/COPQ_Quality_Audit_Report.json`

### 3. Run Pipeline Orchestration Atomically
```bash
python agents/run_agent.py orchestrate --period "2026-08"
```
Generated artifacts:
- `Output/COPQ_Clean.xlsx`
- `Output/FTT_Combined_Report.xlsx`
- `Output/COPQ_Report_<YYYY-MM>.pptx`
- `Output/copq_pipeline_telemetry.json`

### 4. Generate C-Suite Executive Intelligence Briefing
```bash
python agents/run_agent.py brief --period "Aug, 2026"
```
Generated artifacts:
- `Output/COPQ_Executive_Briefing.md`
- `Output/COPQ_Executive_Briefing.html`

### 5. Launch Continuous Watcher Daemon
```bash
python agents/run_agent.py watch --poll-interval 15
```

---

## 🧠 Diagnostic & Triage Runbook

### When Health Score < 90 / Action Required:
1. **Missing Factory Sites**:
   - Check `COPQ_Input/` for files from `VH`, `VH2`, `VH3`, `VH4`, `JV`, `JV2`, `JV3`.
   - If files have non-standard names, verify `extract_site_name()` in `pipeline_common.py`.
2. **Unmapped Defect Terms**:
   - Check `Output/COPQ_Quality_Audit_Report.md` section 3.
   - If new defects are found, add their corresponding HEX/RGB color definitions to `BC_Color/Color_Defect.xlsx`.
3. **Locked Output Files (Excel / PPTX Open in Windows)**:
   - The agents leverage `save_with_fallback()` to save to `_v1`, `_v2`, etc. if files are locked in Microsoft Office.
   - Close open Office instances and rerun `orchestrate` to write cleanly to primary paths.
