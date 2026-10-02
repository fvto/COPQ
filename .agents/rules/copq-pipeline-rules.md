# COPQ Manufacturing Analytics & Agent Automation Rules

## 1. SOP Defect Classification Rules
- Classification into standard categories (`Touch-up`, `Reinspection`, `B/C`, `Rework`) must strictly adhere to SOP heuristics:
  - If keyword `"touch"` in `Type` or `"touch-up"` / `"touch up"` in `Remark (Re-inspection Station)` -> Classify as `"Touch-up"`.
  - If `Type` is `"b/c"` or `"bc"` -> Classify as `"B/C"`.
  - If `"reinspect"` or `"re-inspection"` in `Type` -> Classify as `"Reinspection"`.
  - If `"rework"` in `Type` -> Classify as `"Rework"`.
  - Otherwise -> Classify as `"Other"`.

## 2. Multi-Site Integrity & Alias Handling
- The 7 manufacturing sites are:
  - Vietnam: `VH`, `VH2`, `VH3`, `VH4`
  - Indonesia: `JV`, `JV2`, `JV3` (alias: `JVB`)
- `JVB` and `JV3` refer to the same plant facility. When mapping to executive presentation decks or pivot tables, ensure the alias mapping `JV3 <-> JVB` is uniformly maintained.

## 3. Safe File I/O & Locked Workbook Policy
- Production workbooks and PowerPoint presentations (`COPQ_Clean.xlsx`, `FTT_Combined_Report.xlsx`, `COPQ_Report_*.pptx`) may be opened by end users in Microsoft Excel or PowerPoint.
- Any script writing output must employ `save_with_fallback()` to handle `PermissionError` cleanly without crashing the pipeline.

## 4. Proactive Execution Standard
- Do not rely solely on manual UI button clicks. All pipelines must be executable via the unified agent runner (`python agents/run_agent.py`).
- Always run the Quality Sentinel (`python agents/run_agent.py audit`) when incoming data formats change or when new factory sites are onboarded.
