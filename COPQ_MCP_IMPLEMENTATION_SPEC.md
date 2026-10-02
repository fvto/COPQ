# COPQ Pipeline & Analytics Toolkit — MCP Implementation Specification

## 1. Scope, readiness, and evidence status

### 1.1 Scope of Handoff
This specification provides the complete, authoritative technical handoff for the **COPQ Pipeline & Analytics Toolkit** (`combine_COPQ`). It specifies how to integrate the standalone automated manufacturing analytics pipeline—covering Cost of Poor Quality (COPQ) and First-Time-Through (FTT) data cleaning, SOP defect classification, dynamic Excel formula/chart generation, database synchronization, and executive presentation assembly—into an MCP-compliant architecture (such as `pc_tool_agent`).

### 1.2 Readiness Classification
**READY_FOR_IMPLEMENTATION**

All locally authored algorithms, site normalization rules, Standard Operating Procedure (SOP) classification functions, Excel formula generators, openpyxl charting routines, database cloning logic, and presentation update templates are present and verified in the bundled standalone source tree. Essential static binary assets (`BC_Color/Color_Defect.xlsx`, `Database/CoPQ database 25.xlsx`, `Database/CoPQ_type_analysis.xlsx`, `Template_COPQ/Template_COPQ.pptx`) are preserved verbatim in this package. Synthetic test fixtures are included to allow full, isolated receiver verification without accessing the extractor's machine or private data.

### 1.3 Mode Selection
**Mode B: `copq_pipeline_MCP_HANDOFF.zip`**
- **Rationale**: The project critically depends on four binary Microsoft Office files (`BC_Color/Color_Defect.xlsx`, `Database/CoPQ database 25.xlsx`, `Database/CoPQ_type_analysis.xlsx`, and `Template_COPQ/Template_COPQ.pptx`). These files contain cell-level color fills, multi-sheet formulas, embedded XML chart definitions, and PowerPoint slide shape layouts that cannot be represented in plain Markdown text without risk of formula truncation or binary corruption. Mode B bundles these assets alongside the Markdown specification and synthetic fixtures in a single portable archive.

### 1.4 Evidence Standards
Every statement in this document is verified against repository artifacts and marked with one of the following evidence tags:
- **CONFIRMED_BY_CODE**: Verified directly by reading active source code files and functions.
- **CONFIRMED_BY_CONFIG**: Verified in repository configuration or schema definitions.
- **CONFIRMED_BY_TEST**: Verified by executing automated tests and parity scripts.
- **CONFIRMED_BY_RUNTIME**: Verified through observed execution logs and runtime telemetry.
- **CONFIRMED_BY_USER**: Provided by explicit user instructions.
- **INFERRED**: Deduced through logical analysis of multiple matching code behaviors.
- **UNKNOWN**: Information not present in the standalone repository; marked explicitly.

---

## 2. Source and target project identities

### 2.1 Source Project Identity
- **Project Root**: `c:\Users\User\Desktop\combine_COPQ` [CONFIRMED_BY_CODE]
- **Project Name / Canonical Identifier**: `combine_COPQ` / `copq_pipeline` [CONFIRMED_BY_CODE]
- **Workflow Version**: 2.0.0 (Autonomous Orchestrator & Sentinel Architecture) [CONFIRMED_BY_CONFIG: `agents/pipeline_orchestrator.py:2`]
- **Extraction Timestamp**: 2026-09-29T08:26:00+07:00 [CONFIRMED_BY_RUNTIME]
- **Primary Source Modules**: `combine_files.py`, `process_ftt.py`, `monthly_reporting.py`, `pipeline_common.py`, `Output/write_db_month.py`, `Output/update_copq_pptx.py`, `agents/pipeline_orchestrator.py`, `agents/quality_sentinel.py`, `agents/watchdog_delivery.py`, `agents/run_agent.py`.

### 2.2 Target Project Identity
- **Target MCP Project Root**: `[TARGET_MCP_ROOT_OR_UNAVAILABLE]` -> **UNAVAILABLE** [CONFIRMED_BY_USER]
- **Architecture Baseline**: The target architecture is assumed to be an MCP server (`pc_tool_agent`) providing Tool Registration, Input Validation, Security Sandboxing, and Structured JSON Telemetry Output.
- **Porting Assumption**: AI #2 will implement the MCP tool within the target server without modifying the verified data-processing algorithms.

---

## 3. Entry points and stage dependency graph

### 3.1 Executable Entry Points
| Entry point | Trigger | Calls | Inputs supplied | Working directory | Browser/network use | Side effects | Is this the requested workflow? |
|---|---|---|---|---|---|---|---|
| `agents/run_agent.py orchestrate` | CLI command | `PipelineOrchestratorAgent.execute_pipeline()` | `--period`, `--copq-dir`, `--ftt-dir`, `--output-dir` | Standalone root | None (0 network calls) | Writes 4 output artifacts, updates DB, generates telemetry | **Yes** (Primary Orchestrated Workflow) [CONFIRMED_BY_CODE: `agents/run_agent.py:42-79`] |
| `agents/run_agent.py audit` | CLI command | `QualitySentinelAgent.run_full_audit()` | `--period` | Standalone root | None | Generates JSON/MD quality audit reports | Companion Quality Sentinel Workflow [CONFIRMED_BY_CODE: `agents/run_agent.py:49-85`] |
| `agents/run_agent.py brief` | CLI command | `WatchdogDeliveryAgent.generate_executive_briefing()` | `--period` | Standalone root | None | Generates MD and HTML briefing decks | Companion Executive Reporting Workflow [CONFIRMED_BY_CODE: `agents/run_agent.py:53-90`] |
| `combine_files.py` | CLI / Script | `combine_copq_files()` -> `build_copq_clean_master()` | `--input <DIR>`, `--output <FILE>` | Standalone root | None | Reads raw ERP files, writes `Output/COPQ_Clean.xlsx` | **Yes** (Stage 1 Core Processor) [CONFIRMED_BY_CODE: `combine_files.py:270-340`] |
| `process_ftt.py` | CLI / Script | `generate_ftt_report()` | `--ftt-dir`, `--output`, `--color-map`, `--cost-file` | Standalone root | None | Reads FTT files, writes `Output/FTT_Combined_Report.xlsx` | **Yes** (Stage 2 Core Processor) [CONFIRMED_BY_CODE: `process_ftt.py:650-750`] |
| `Output/write_db_month.py` | CLI / Script | `main()` -> `parse_from_clean_df()` | Implicit paths | Standalone root or `Output/` | None | Reads `COPQ_Clean.xlsx`, clones/updates DB workbooks | **Yes** (Stage 3 Database Sync) [CONFIRMED_BY_CODE: `Output/write_db_month.py:680-709`] |
| `Output/update_copq_pptx.py` | CLI / Script | `main()` -> `update_presentation()` | Implicit paths | Standalone root or `Output/` | None | Reads DB & Excel outputs, writes `COPQ_Report_<period>.pptx` | **Yes** (Stage 4 Slide Generator) [CONFIRMED_BY_CODE: `Output/update_copq_pptx.py:1800-1860`] |
| `Interactive_Dashboard.py` | GUI launch | Tkinter desktop UI embedding Matplotlib | Button clicks | Standalone root | None | Spawns child scripts via `sys.executable` | **No** (Desktop UI to be omitted in MCP) [CONFIRMED_BY_CODE] |
| `Run_Interactive_Dashboard.bat`, `process_ftt.bat` | Shell launch | Batch script launching python | None | Standalone root | None | Environment checks, pip install, script launch | **No** (OS shell scripts to be omitted in MCP) [CONFIRMED_BY_CODE] |

### 3.2 Stage Dependency Call Graph
```text
[COPQ ERP Exports in COPQ_Input/] ───┐
                                      ├──> [Stage 1: combine_files.py] ───> Output/COPQ_Clean.xlsx
                                      │                                             │
[FTT ERP Exports in FTT_Input/]   ───┼───> [Stage 2: process_ftt.py]    ───> Output/FTT_Combined_Report.xlsx
                                      │           ▲                                 │
[BC_Color/Color_Defect.xlsx]      ───┤           │                                 │
                                      │           ▼                                 ▼
[Database/CoPQ_type_analysis.xlsx] ──┴──> [Stage 3: write_db_month.py]  ───> Cloned/Updated Database Files
                                                  │                                 │
[Template_COPQ/Template_COPQ.pptx] ───────────────┴──> [Stage 4: update_copq_pptx] ─┴──> Output/COPQ_Report_<Period>.pptx
```

---

## 4. Successful runtime flow

### 4.1 Stage Matrix
| Stage | Prerequisites | Exact function/action | Output or state | Success evidence | Failure/partial behavior | Retry unit |
|---|---|---|---|---|---|---|
| **1. COPQ Ingestion & Consolidation** | Raw COPQ export files exist in input folder | `combine_files.combine_copq_files()` | `Output/COPQ_Clean.xlsx` created with per-site sheets + `COPQ_Clean` master | File exists, size > 10KB, sheet `COPQ_Clean` has rows | Raises `FileNotFoundError` or `ValueError` on 0 valid files | Rerun Stage 1 with verified input directory [CONFIRMED_BY_CODE: `combine_files.py:310`] |
| **2. FTT Defect Processing & Charting** | FTT export files exist; `Color_Defect.xlsx` exists; Stage 1 complete | `process_ftt.generate_ftt_report()` | `Output/FTT_Combined_Report.xlsx` with `Top Defects Summary`, dynamic formulas, bar charts | File exists, size > 10KB, sheet `Top Defects Summary` contains `=SUMIFS` | Falls back to Qty ranking if cost sheet missing; raises if 0 files | Rerun Stage 2 [CONFIRMED_BY_CODE: `process_ftt.py:680`] |
| **3. Database Synchronization** | `COPQ_Clean.xlsx` exists; Database master files exist | `write_db_month.main()` / `monthly_reporting.create_monthly_database_clones()` | `Output/Monthly/<Period>/Database/` created with new month sheet; DB master updated | Target sheets exist (e.g. `Aug-26`), overview sums computed | Aborts if clean sheet missing or date unrecognized | Rerun Stage 3 [CONFIRMED_BY_CODE: `monthly_reporting.py:120`] |
| **4. Executive Presentation Generation** | PPTX template exists; DB overview and clean data exist | `update_copq_pptx.main()` | `Output/COPQ_Report_<Period>.pptx` generated with updated shapes and tables | PPTX file exists, size > 100KB, slide 1 titles match period | Logs warning on missing metric; keeps template value | Rerun Stage 4 [CONFIRMED_BY_CODE: `Output/update_copq_pptx.py:1820`] |

### 4.2 Multi-Site Processing Order
The standalone system strictly processes seven factory site codes in a standardized order [CONFIRMED_BY_CODE: `pipeline_common.py:14-19`]:
1. `VH` (Factory 1, Vietnam)
2. `VH2` (Factory 2, Vietnam)
3. `VH3` (Factory 3, Vietnam)
4. `VH4` (Factory 4, Vietnam)
5. `JV` (Factory 1, Indonesia)
6. `JV2` (Factory 2, Indonesia)
7. `JV3` (Factory 3, Indonesia; normalized from `JVB` where applicable)

In Stage 1 and Stage 2, if an input file for a particular factory is missing, the processor logs a warning and proceeds with the available factories [CONFIRMED_BY_CODE: `combine_files.py:285`]. The resulting workbook contains sheets only for sites with valid data, while `QualitySentinelAgent` flags missing sites as warnings in telemetry [CONFIRMED_BY_RUNTIME].

---

## 5. Input contract and data schema crosswalk

### 5.1 Raw COPQ Export Contract (`COPQ_Input/`)
- **File Format**: Excel `.xlsx`, `.xls` (BIFF8), or `.csv` [CONFIRMED_BY_CODE: `combine_files.py:18-20`].
- **Filename Convention**: Must contain date range and site code, e.g., `Cost Of Poor Quality  2026-08-01_2026-08-31-VH.xls` [CONFIRMED_BY_CODE: `pipeline_common.py:35`].
- **Header Structure**: Raw ERP exports contain multi-row banner headers (e.g. title banners, empty spacer rows). The processor dynamically scans rows 0 to 9 to find the row containing both `"date"` and either `"factory"` or `"fty"` [CONFIRMED_BY_CODE: `combine_files.py:45-56`].
- **Canonical Schema (20 Columns)**:
  `Date`, `Category`, `Type`, `Factory`, `Fty/Line`, `Style Nbr`, `Mold`, `Model`, `Size`, `Part`, `Defect Name`, `Sub Defect Name`, `Cause Area`, `Station`, `Process`, `Issue`, `Qty`, `Cost`, `Labor Cost`, `Material Cost`, `Remark (Re-inspection Station)`.
- **Spacer Columns**: Unnamed columns with all null values or header matching `^unnamed` are automatically purged [CONFIRMED_BY_CODE: `combine_files.py:72`].

### 5.2 Raw FTT Defect Export Contract (`FTT_Input/`)
- **File Format**: Excel `.xlsx`, `.xls` [CONFIRMED_BY_CODE: `process_ftt.py:25`].
- **Filename Convention**: e.g., `FTT-Internal_2026-08-01_2026-08-31-VH.xlsx`.
- **Model2 Priority Matching**: Columns are mapped via regex. If `Model 2` or `Model2` exists, it is prioritized over generic `Model` because `Model 2` contains commercial production footwear silhouette names [CONFIRMED_BY_CODE: `process_ftt.py:95-115`].
- **Grade B/C Filtering**: Only rows where `Metric` or `Grade` matches `(B/C|Grade B|Grade C|B grade|C grade)` are retained [CONFIRMED_BY_CODE: `process_ftt.py:180-195`].

### 5.3 Static Reference Assets Contract
1. `BC_Color/Color_Defect.xlsx`:
   - Contains sheets `ColorMap_bc` and `ColorMap_BC New 2026` [CONFIRMED_BY_CODE: `process_ftt.py:220`].
   - Schema: `Defect Name` (string), `HEX` (6-char string e.g. `FFC000`) or separate `R`, `G`, `B` columns.
2. `Database/CoPQ database 25.xlsx`:
   - Master historical database containing monthly sheets (`Nov-24`, `Dec-24`, ..., `Current`).
3. `Database/CoPQ_type_analysis.xlsx`:
   - Master historical database containing `Current month` sheet with B/C cost ranking blocks:
     Headers: `FTY` | `Model` | `Defect Qty` | `Ttl cost ($)` [CONFIRMED_BY_CODE: `process_ftt.py:280-310`].
4. `Template_COPQ/Template_COPQ.pptx`:
   - Base executive PowerPoint template containing pre-formatted shapes, tables, and native PowerPoint charts.

### 5.4 Schema Crosswalk Table
| Upstream Raw ERP Column | Canonical COPQ Standard Column | Data Type | Fallback / Normalization Rule |
|---|---|---|---|
| `Date` / `Ngay` | `Date` | `string` (`YYYY-MM-DD`) | Parsed via `pd.to_datetime`; coerced to ISO date string [CONFIRMED_BY_CODE: `combine_files.py:85`] |
| `Category` | `Category` | `string` | Stripped; defaults to `Internal` if blank |
| `Type` | `Type` | `string` | Stripped; classified via `classify_copq_type()` |
| `Factory` / `Xuong` | `Factory` | `string` | Stripped |
| `Fty/Line` / `Chuyen` | `Fty/Line` | `string` | Stripped |
| `Style Nbr` | `Style Nbr` | `string` | Stripped; preserves alphanumeric style codes |
| `Model 2` / `Model2` | `Model` (in FTT) | `string` | Takes precedence over `Model` [CONFIRMED_BY_CODE: `process_ftt.py:102`] |
| `Defective Qty(Pair)` / `Qty` | `Qty` | `float` / `int` | Cleaned of commas, cast to float, defaults to 0.0 |
| `Ttl cost ($)` / `Cost` | `Cost` | `float` | Cleaned of currency symbols, cast to float, defaults to 0.0 |
| `Remark (Re-inspection Station)` | `Remark (Re-inspection Station)` | `string` | Critical override: if contains "touch-up" -> classifies as Touch-up [CONFIRMED_BY_CODE: `monthly_reporting.py:48`] |

---

## 6. Acquisition and upstream/downstream handoff

### 6.1 Data Acquisition Flow
- In the standalone deployment, raw ERP exports are extracted manually from enterprise SAP/ERP portals or scheduled network shares and placed directly into `COPQ_Input/` and `FTT_Input/` [CONFIRMED_BY_RUNTIME].
- **No Browser Companion**: The standalone project contains zero Selenium/Playwright browser automation. The supported input contract is strictly local directory ingestion. The MCP tool should expose explicit directory/file parameters rather than attempting browser emulation [CONFIRMED_BY_CODE].

### 6.2 Upstream to Downstream Cardinality
- **COPQ Files**: 1 to 7 files per monthly period (one per factory site: `VH`, `VH2`, `VH3`, `VH4`, `JV`, `JV2`, `JV3`).
- **FTT Files**: 1 to 4 files per monthly period (`VH`, `VH2`, `JV`, `JV2`).
- **Downstream Artifacts**: Exactly 1 `COPQ_Clean.xlsx`, 1 `FTT_Combined_Report.xlsx`, 1 monthly cloned database directory, and 1 `COPQ_Report_<Period>.pptx`.

---

## 7. Business rules and calculations

### 7.1 SOP Defect Type Classification Logic
Implemented in `monthly_reporting.classify_copq_type(row)` [CONFIRMED_BY_CODE: `monthly_reporting.py:35-55`]:
```python
def classify_copq_type(row):
    t = str(row.get("Type", "")).strip().lower()
    remark = str(row.get("Remark (Re-inspection Station)", "")).strip().lower()

    # Rule 1: Touch-up priority (including remark override)
    if "touch" in t or "touch-up" in remark or "touch up" in remark:
        return "Touch-up"
    # Rule 2: Grade B/C defect
    if t in ("b/c", "bc"):
        return "B/C"
    # Rule 3: Reinspection
    if "reinspect" in t or "re-inspection" in t:
        return "Reinspection"
    # Rule 4: Rework
    if "rework" in t:
        return "Rework"
    # Rule 5: Fallback
    return "Other"
```
> [!IMPORTANT]
> The remark override rule is essential: a record whose `Type` is "Reinspection" MUST be classified as "Touch-up" if its `Remark (Re-inspection Station)` field contains "touch-up" or "touch up".

### 7.2 Site Name Extraction & Normalization
Implemented in `pipeline_common.extract_site_name(filename)` [CONFIRMED_BY_CODE: `pipeline_common.py:22-42`]:
- Candidate evaluation order: `["VH2", "VH3", "VH4", "JV2", "JV3", "JVB", "VH", "JV"]`.
- `VH2` must be evaluated before `VH`; `JV2` before `JV`.
- If match is `JVB`, it is normalized to `JV3`.

### 7.3 Top 3 Models & Defect Matrix Generation
Implemented in `process_ftt.write_site_matrix()` [CONFIRMED_BY_CODE: `process_ftt.py:380-450`]:
1. Top 3 models by site are selected based on total cost from `Database/CoPQ_type_analysis.xlsx` (`Current month` sheet).
2. If cost data is missing or site not found, falls back to Top 3 models by defect quantity from the FTT data.
3. For each model, the Top 3 defect issues are identified.
4. Dynamic Excel formulas are inserted into the matrix:
   ```text
   =IFERROR(SUMIFS('Site_VH'!$O:$O, 'Site_VH'!$C:$C, $B5, 'Site_VH'!$M:$M, C$4), 0)
   ```
5. Defect Rate (DR%) is calculated dynamically:
   ```text
   =IFERROR(Defect_Qty / Total_Produced, 0)
   ```
6. Openpyxl `BarChart` (stacked column) is created, referencing the matrix cells. Exact HEX colors from `Color_Defect.xlsx` are applied to chart series XML tags.

---

## 8. File, template, database, and temporary lifecycle

### 8.1 File Lifecycle Table
| Path/pattern | Stage | Origin | Read/write/move/delete | Fixed or user-controlled | Collision behavior | Retained on success | Retained on failure | Evidence |
|---|---|---|---|---|---|---|---|---|
| `COPQ_Input/*.xls*` | Stage 1 | ERP Export drop | Read-only | User-controlled | None (read only) | Yes | Yes | `combine_files.py:270` |
| `FTT_Input/*.xlsx` | Stage 2 | ERP Export drop | Read-only | User-controlled | None (read only) | Yes | Yes | `process_ftt.py:650` |
| `Output/COPQ_Clean.xlsx` | Stage 1 | Script generated | Write / Overwrite | Standard fixed | `save_with_fallback()` (`_vN`) | Yes | Partial file retained | `pipeline_common.py:45` |
| `Output/FTT_Combined_Report.xlsx` | Stage 2 | Script generated | Write / Overwrite | Standard fixed | `save_with_fallback()` (`_vN`) | Yes | Partial file retained | `pipeline_common.py:45` |
| `Database/CoPQ database 25.xlsx` | Stage 3 | Master asset | Read & In-place update | Standard fixed | Direct cell modification | Yes | Yes (never deleted) | `write_db_month.py:400` |
| `Output/Monthly/<Period>/Database/` | Stage 3 | Cloned from master | Write (cloned copies) | Standard fixed | Overwrites existing month dir | Yes | Yes | `monthly_reporting.py:115` |
| `Template_COPQ/Template_COPQ.pptx` | Stage 4 | Master asset | Read-only template | Standard fixed | Never modified in-place | Yes | Yes | `update_copq_pptx.py:10` |
| `Output/COPQ_Report_<Period>.pptx` | Stage 4 | Script generated | Write / Overwrite | Standard fixed | Overwrite target | Yes | Partial file retained | `update_copq_pptx.py:1840` |

### 8.2 Locked-File Fallback Mechanism
Implemented in `pipeline_common.save_with_fallback(book, output_path, max_retries=19)` [CONFIRMED_BY_CODE: `pipeline_common.py:45-72`]:
- If the target file is open in Microsoft Excel or PowerPoint, saving triggers a `PermissionError`.
- The function intercepts the exception and attempts saving to `<stem>_v1.<ext>`, `<stem>_v2.<ext>`, ..., up to `<stem>_v19.<ext>`.
- The final saved path is logged and returned to the caller.

---

## 9. Runtime dependencies and asset provisioning

### 9.1 Runtime Dependencies Matrix
| Requirement | Category | Required for which stage | Version or format | Provisioning source | Resolution base | Runtime location | Missing behavior | Evidence |
|---|---|---|---|---|---|---|---|---|
| `python` | System Runtime | All stages | Python 3.9+ | System / Virtualenv | PATH | System PATH | Fatal error | Runtime preflight |
| `pandas` | Python Package | Stages 1, 2, 3, 4 | >= 1.5.0 | PyPI (`pip install pandas`) | Python site-packages | Python env | `ImportError` | `combine_files.py:1` |
| `openpyxl` | Python Package | Stages 1, 2, 3 | >= 3.0.0 | PyPI (`pip install openpyxl`) | Python site-packages | Python env | `ImportError` | `combine_files.py:2` |
| `numpy` | Python Package | Stages 1, 2 | >= 1.20.0 | PyPI (`pip install numpy`) | Python site-packages | Python env | `ImportError` | `process_ftt.py:8` |
| `python-pptx` | Python Package | Stage 4 | >= 0.6.21 | PyPI (`pip install python-pptx`) | Python site-packages | Python env | `ImportError` | `update_copq_pptx.py:39` |
| `lxml` | Python Package | Stage 4 | >= 4.9.0 | PyPI (`pip install lxml`) | Python site-packages | Python env | `ImportError` | `update_copq_pptx.py:35` |
| `Pillow` (`PIL`) | Python Package | Stage 4 | >= 9.0.0 | PyPI (`pip install Pillow`) | Python site-packages | Python env | `ImportError` | `update_copq_pptx.py:36` |
| `Color_Defect.xlsx` | Static Binary Asset | Stage 2 | Excel Workbook | Bundled in Mode B handoff | Repo root | `BC_Color/Color_Defect.xlsx` | FTT uses fallback colors | `process_ftt.py:220` |
| `CoPQ database 25.xlsx` | Master Database | Stage 3, 4 | Excel Workbook | Bundled in Mode B handoff | Repo root | `Database/CoPQ database 25.xlsx` | Stage 3 fails | `write_db_month.py:23` |
| `CoPQ_type_analysis.xlsx` | Master Database | Stage 2, 3 | Excel Workbook | Bundled in Mode B handoff | Repo root | `Database/CoPQ_type_analysis.xlsx` | FTT uses Qty fallback | `process_ftt.py:280` |
| `Template_COPQ.pptx` | Presentation Template | Stage 4 | PowerPoint Presentation | Bundled in Mode B handoff | Repo root | `Template_COPQ/Template_COPQ.pptx` | Stage 4 fails | `update_copq_pptx.py:65` |

---

## 10. Configuration, paths, and deployment portability

### 10.1 Root-Relative Path Resolution
All file operations in the standalone code resolve paths dynamically relative to `ROOT_DIR = os.path.dirname(...)` [CONFIRMED_BY_CODE: `agents/pipeline_orchestrator.py:29`]. There are zero hardcoded developer machine paths (`C:\Users\User\...`) in core logic functions.

### 10.2 Recommended MCP Sandboxing
The target MCP server must enforce directory path containment:
- All input paths (`copq_input_dir`, `ftt_input_dir`) and output paths (`output_dir`) must be verified to reside within approved workspace boundaries.
- Symlinks resolving outside the root directory must be rejected.

---

## 11. Browser, network, and subprocess behavior

### 11.1 Network and Browser Audit
- **Network Calls**: **Zero (0)**. The pipeline performs no HTTP requests, no database socket connections, and no cloud uploads [CONFIRMED_BY_CODE].
- **Browser Automation**: **Zero (0)**. No Selenium, Playwright, or CDP connections are used [CONFIRMED_BY_CODE].

### 11.2 Subprocess Elimination
- In the desktop standalone, `Interactive_Dashboard.py` spawned helper scripts via `subprocess.run([sys.executable, script_path])`.
- In the MCP adaptation, all four pipeline stages MUST be executed **in-process** via direct Python module imports (`combine_files.combine_copq_files()`, `process_ftt.generate_ftt_report()`, etc.), eliminating subprocess overhead and platform-specific shell quirks [CONFIRMED_BY_CODE: `agents/pipeline_orchestrator.py:42-45`].

---

## 12. Output and delivery contract

### 12.1 Produced Artifacts
1. `Output/COPQ_Clean.xlsx`:
   - Worksheets: Per-site sheets (`VH`, `VH2`, `VH3`, `VH4`, `JV`, `JV2`, `JV3`) preserving the original 2-row header structure; `COPQ_Clean` consolidated master sheet with `Site` column prepended; pivot analysis sheets (`COPQ_Pivot_Overview`, `COPQ_Pivot_TopFactory`, `COPQ_Pivot_TopModel`, `COPQ_Report`) [CONFIRMED_BY_CODE: `combine_files.py:330`].
2. `Output/FTT_Combined_Report.xlsx`:
   - Worksheets: `Combined_All` (raw cleaned B/C records across all sites); per-site sheets (`Site_VH`, `Site_VH2`, `Site_JV`, `Site_JV2`); embedded color maps (`ColorMap_bc`, `ColorMap_BC New 2026`); `Top Defects Summary` executive dashboard with `=SUMIFS` formulas and stacked column charts [CONFIRMED_BY_CODE: `process_ftt.py:680-730`].
3. Monthly Database Clones:
   - Output location: `Output/Monthly/<YYYY-MM>/Database/` containing cloned copies of `CoPQ database 25.xlsx` and `CoPQ_type_analysis.xlsx` with newly populated monthly sheets [CONFIRMED_BY_CODE: `monthly_reporting.py:115`].
4. `Output/COPQ_Report_<Period>.pptx`:
   - Populated executive slide deck updating period strings, overview summary KPI shapes, defect share charts, and factory callouts [CONFIRMED_BY_CODE: `Output/update_copq_pptx.py:1840`].
5. Telemetry & Audit Artifacts:
   - `Output/copq_pipeline_telemetry.json`: Execution metadata, timing, and status.
   - `Output/COPQ_Quality_Audit_Report.json` & `.md`: Data quality audit findings.

---

## 13. Terminal success predicate and partial/failure semantics

### 13.1 Terminal Success Predicate
A workflow execution is deemed `SUCCESS` if and only if all eight of the following conditions evaluate to `true` [CONFIRMED_BY_CODE: `agents/pipeline_orchestrator.py:220-250`]:
1. `os.path.exists(copq_clean_path)` AND `os.path.getsize(copq_clean_path) > 10240`.
2. `openpyxl.load_workbook(copq_clean_path)` loads without exception AND sheet `COPQ_Clean` has >= 1 data row.
3. `os.path.exists(ftt_report_path)` AND `os.path.getsize(ftt_report_path) > 10240`.
4. `openpyxl.load_workbook(ftt_report_path)` contains sheet `Top Defects Summary` with >= 1 valid `=SUMIFS` formula string.
5. Monthly archive folder `Output/Monthly/<Period>/Database/` exists AND contains both database workbooks.
6. Target month sheet (e.g. `Aug-26`) exists in the updated `CoPQ database 25.xlsx`.
7. `os.path.exists(pptx_path)` AND `os.path.getsize(pptx_path) > 102400`.
8. `Presentation(pptx_path)` loads without exception AND slide 1 title text contains `<Period_Label>`.

### 13.2 Status Codes
- `SUCCESS`: All 4 stages completed and satisfied the terminal predicate.
- `PARTIAL`: Stages 1 and 2 succeeded, but Stage 3 or 4 failed or was skipped due to missing template/database.
- `FAILED`: Stage 1 or 2 failed to produce valid workbooks.
- `CANCELLED`: Interrupted by caller before stage completion.

---

## 14. State, concurrency, cleanup, and idempotency

### 14.1 Concurrency and Idempotency
- **Stateless Execution**: The pipeline maintains no long-lived daemon state or database locks.
- **Idempotency**: Rerunning the pipeline with identical inputs safely overwrites prior outputs or writes to `_vN` files if locked.
- **Memory Cleanup**: Workbooks and DataFrames are explicitly closed and dereferenced upon function completion.

---

## 15. Security and confidentiality

### 15.1 Path Traversal Guards
All caller-supplied paths must be sanitized using `os.path.abspath()` and verified to reside within approved directory roots before file operations begin.

### 15.2 Confidentiality and PII
Raw ERP logs may contain operator employee names, line numbers, or proprietary style numbers. Logs and structured MCP responses MUST NOT leak individual worker identities or raw customer pricing data. Only aggregated defect quantities, site codes, and summary costs should be returned in tool output [CONFIRMED_BY_CONFIG].

---

## 16. Target MCP integration matrix

| Standalone Stage / Contract | Target MCP Architecture Role | Exact Adaptation | Config / Path Rule | Error / Status Mapping |
|---|---|---|---|---|
| Stage 1: `combine_copq_files()` | `copq_pipeline` (action: `combine_copq` or `run_all`) | In-process function call | Param: `copq_input_dir`, `output_dir` | Map `ValueError` / `FileNotFoundError` to MCP tool error |
| Stage 2: `generate_ftt_report()` | `copq_pipeline` (action: `process_ftt` or `run_all`) | In-process function call | Param: `ftt_input_dir`, `color_map_file`, `cost_file` | Log warning on missing cost file; return fallback |
| Stage 3: `write_db_month.py` | `copq_pipeline` (action: `update_database` or `run_all`) | In-process function call | Param: `database_dir`, `output_dir` | Return `PARTIAL` if DB sheets cannot be updated |
| Stage 4: `update_copq_pptx.py` | `copq_pipeline` (action: `generate_presentation` or `run_all`) | In-process function call | Param: `template_pptx`, `output_dir` | Log warning if slide shape missing; keep existing |
| Quality Sentinel: `QualitySentinelAgent` | `copq_pipeline` (action: `audit_quality`) | In-process audit call | Param: `copq_input_dir`, `ftt_input_dir` | Return health score and anomaly list |

---

## 17. Proposed tool input and structured output schemas

### 17.1 Input Schema (JSON Schema)
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "CopqPipelineInput",
  "type": "object",
  "properties": {
    "action": {
      "type": "string",
      "enum": ["run_all", "combine_copq", "process_ftt", "update_database", "generate_presentation", "audit_quality"],
      "default": "run_all",
      "description": "Pipeline stage or workflow to execute."
    },
    "period": {
      "type": "string",
      "description": "Target reporting period (e.g. '2026-08' or 'Aug, 2026'). If omitted, auto-detected from filenames."
    },
    "copq_input_dir": {
      "type": "string",
      "description": "Path to directory containing raw COPQ ERP export files."
    },
    "ftt_input_dir": {
      "type": "string",
      "description": "Path to directory containing raw FTT defect export files."
    },
    "output_dir": {
      "type": "string",
      "description": "Directory where generated Excel, PPTX, and telemetry files will be saved."
    },
    "color_map_file": {
      "type": "string",
      "description": "Path to Color_Defect.xlsx reference workbook."
    },
    "database_dir": {
      "type": "string",
      "description": "Directory containing master database workbooks."
    },
    "template_pptx": {
      "type": "string",
      "description": "Path to executive PowerPoint template file."
    },
    "update_master_database": {
      "type": "boolean",
      "default": false,
      "description": "If true, updates master database files in-place in addition to creating monthly clones."
    }
  },
  "required": ["action"]
}
```

### 17.2 Structured Output Schema (JSON Schema)
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "CopqPipelineOutput",
  "type": "object",
  "properties": {
    "status": {
      "type": "string",
      "enum": ["SUCCESS", "PARTIAL", "FAILED", "CANCELLED"]
    },
    "period": {
      "type": "object",
      "properties": {
        "iso": { "type": "string" },
        "label": { "type": "string" },
        "db_month": { "type": "string" }
      }
    },
    "execution_duration_seconds": { "type": "number" },
    "stages_completed": {
      "type": "array",
      "items": { "type": "string" }
    },
    "generated_artifacts": {
      "type": "object",
      "properties": {
        "copq_clean": { "type": "string" },
        "ftt_report": { "type": "string" },
        "monthly_database_dir": { "type": "string" },
        "pptx_presentation": { "type": "string" },
        "telemetry_file": { "type": "string" }
      }
    },
    "warnings": {
      "type": "array",
      "items": { "type": "string" }
    },
    "errors": {
      "type": "array",
      "items": { "type": "string" }
    }
  },
  "required": ["status", "stages_completed", "generated_artifacts"]
}
```

---

## 18. Gateway and caller responsibilities

### 18.1 Gateway Responsibilities
1. **Validation**: Check existence of input directories and required static reference assets before launching pipeline stages.
2. **Sandbox Boundary Enforcement**: Ensure all output files are confined to the designated workspace output directory.
3. **Execution Timeout**: Enforce a default timeout of 180 seconds for the full 4-stage pipeline run.
4. **Telemetry Export**: Automatically write `copq_pipeline_telemetry.json` capturing execution duration and artifact sizes.

---

## 19. Test and verification matrix

### 19.1 Test Scenarios
1. **Contract Test (Headers & Schemas)**:
   - Ingest `Cost Of Poor Quality  2026-08-01_2026-08-31-VH.xlsx`. Verify header is detected on row 2 (0-indexed). Verify output has exactly 21 columns with `Site` in position 0.
2. **Unit Test (SOP Classification & Touch-up Override)**:
   - Record with `Type = "Reinspection"` and `Remark = "Touch-up paint required"`. Verify `classify_copq_type()` returns `"Touch-up"`.
3. **Unit Test (Site Name Extraction)**:
   - Filenames `...-VH2.xls` -> `VH2`, `...-VH.xls` -> `VH`, `...-JVB.xls` -> `JV3`.
4. **Adapter Test (FTT Model2 Priority)**:
   - Input row with `Model = "AIR MAX"`, `Model 2 = "AIR MAX 90"`. Verify output row uses `"AIR MAX 90"`.
5. **Semantic Integration Test (Matrix & Chart Generation)**:
   - Run Stage 2 on synthetic FTT data. Load `FTT_Combined_Report.xlsx`. Verify cell formulas start with `=IFERROR(SUMIFS(`. Verify chart object is present in worksheet `Top Defects Summary`.
6. **Negative Test (Missing Input Directory)**:
   - Call `combine_copq_files(input_dir="non_existent")`. Verify graceful error returned, not an unhandled crash.
7. **Negative Test (Locked File Fallback)**:
   - Hold a read-lock on `COPQ_Clean.xlsx` during Stage 1. Verify `save_with_fallback()` creates `COPQ_Clean_v1.xlsx` without crashing.

---

## 20. Deployment and runtime preflight

### 20.1 Preflight Checklist
- [ ] Python 3.9+ runtime verified (`python --version`).
- [ ] Dependencies installed: `pandas`, `openpyxl`, `numpy`, `python-pptx`, `lxml`, `Pillow`.
- [ ] Asset verified: `standalone/BC_Color/Color_Defect.xlsx` exists and matches SHA-256 in manifest.
- [ ] Asset verified: `standalone/Database/CoPQ database 25.xlsx` exists and matches SHA-256 in manifest.
- [ ] Asset verified: `standalone/Database/CoPQ_type_analysis.xlsx` exists and matches SHA-256 in manifest.
- [ ] Asset verified: `standalone/Template_COPQ/Template_COPQ.pptx` exists and matches SHA-256 in manifest.
- [ ] Writable `Output/` directory provisioned.

---

## 21. Preserve, adapt, and omit decisions

### 21.1 Decision Table
| Component / Behavior | Action | Rationale |
|---|---|---|
| SOP Classification (`classify_copq_type`) | **PRESERVE** | Core manufacturing business rule; Touch-up remark override is critical. |
| Site extraction regex (`SITE_CODE_CANDIDATES`) | **PRESERVE** | Exact factory candidate precedence must not be altered. |
| Header detection & spacer drop heuristics | **PRESERVE** | ERP exports vary across factories; dynamic scan is mandatory. |
| Formula insertion (`=IFERROR(SUMIFS(...)...)`) | **PRESERVE** | Stakeholder requirement: workbooks must contain dynamic formulas, not static values. |
| Chart series coloring from `Color_Defect.xlsx` | **PRESERVE** | Executive decks require exact standardized defect color fills. |
| Locked-file fallback (`save_with_fallback`) | **PRESERVE** | Eliminates crashes when users leave workbooks open in Excel. |
| CLI / Argparse runners (`combine_files.py`) | **ADAPT** | Wrap with clean Python function signatures suitable for MCP tool handler calls. |
| Subprocess execution via `sys.executable` | **ADAPT** | Execute modules in-process via direct Python imports. |
| Desktop GUI (`Interactive_Dashboard.py`) | **OMIT** | Desktop UI is superseded by the MCP tool protocol. |
| Windows `.bat` launchers | **OMIT** | Shell scripts are specific to local desktop execution. |
| `os.startfile()` invocations | **OMIT** | MCP server must not attempt to launch desktop GUI applications on the host. |

---

## 22. Source-of-truth index

| Requirement / Behavior | Source File and Symbol | Line Numbers |
|---|---|---|
| Factory candidate list & `JVB` -> `JV3` normalization | `pipeline_common.py: SITE_CODE_CANDIDATES`, `extract_site_name()` | Lines 14-42 |
| Safe file saving with `_vN` fallback | `pipeline_common.py: save_with_fallback()` | Lines 45-72 |
| Dynamic header row search & spacer column purge | `combine_files.py: _find_copq_header_index()`, `_drop_spacer_columns()` | Lines 45-75 |
| Standard 20-column schema definition | `combine_files.py: STANDARD_COLUMNS` | Lines 22-38 |
| Master COPQ consolidation | `combine_files.py: build_copq_clean_master()` | Lines 120-160 |
| SOP Defect Type Classification logic | `monthly_reporting.py: classify_copq_type()` | Lines 35-55 |
| Dynamic period detection from filenames/dates | `agents/pipeline_orchestrator.py: detect_period()` | Lines 69-125 |
| Atomic 4-stage pipeline coordination | `agents/pipeline_orchestrator.py: execute_pipeline()` | Lines 130-240 |
| FTT Model2 priority mapping & bilingual header scan | `process_ftt.py: auto_map_ftt_columns()` | Lines 90-135 |
| Grade B/C defect regex filter | `process_ftt.py: process_single_ftt_file()` | Lines 170-205 |
| Defect color map loading & embedding | `process_ftt.py: load_defect_color_map()`, `embed_color_defect_workbook()` | Lines 215-275 |
| Top 3 Model cost map extraction | `process_ftt.py: load_model_cost_map()` | Lines 280-325 |
| Dynamic formula construction & BarChart generation | `process_ftt.py: write_site_matrix()`, `add_top3_chart()` | Lines 380-550 |
| Database clone creation & overview writing | `monthly_reporting.py: create_monthly_database_clones()` | Lines 110-150 |
| In-place database population | `Output/write_db_month.py: parse_from_clean_df()`, `main()` | Lines 54-150, 680-709 |
| PowerPoint slide text replacement preserving formats | `Output/update_copq_pptx.py: _set_text_preserve_format()`, `main()` | Lines 150-195, 1800-1860 |

---

## 23. Assumptions, unknowns, and blockers

### 23.1 Assumptions
1. **Single-Month Processing Assumption**: Raw files processed in a single batch belong to the same monthly period [CONFIRMED_BY_CODE: `monthly_reporting.py:65`].
2. **Factory Site Universe**: Strictly limited to `VH`, `VH2`, `VH3`, `VH4`, `JV`, `JV2`, `JV3` [CONFIRMED_BY_CODE: `pipeline_common.py:14`].
3. **Template Preservation**: Presentation template `Template_COPQ/Template_COPQ.pptx` must never be modified in-place; all runs write to `Output/` [CONFIRMED_BY_CODE: `Output/update_copq_pptx.py:9`].

### 23.2 Unknowns
- Target MCP server repo identity is unavailable in this environment; concrete registration code will be generated by AI #2 using the conceptual schemas in Section 17.

### 23.3 Implementation Blockers
**None**. All algorithms, static binary assets, schemas, and test fixtures are fully provided in this package.

---

## 24. Cross-person transferability and handoff coverage

### 24.1 Transferability Audit
| Required component | Why it is needed | How AI #2 obtains or reproduces it | Pinned version/hash or exact rule | Safe test evidence | Transfer status |
|---|---|---|---|---|---|
| `combine_files.py` | Stage 1 COPQ data cleaning | Bundled in ZIP | `standalone/combine_files.py` | Unit & parity test | **IN_ZIP** |
| `process_ftt.py` | Stage 2 FTT defect processor | Bundled in ZIP | `standalone/process_ftt.py` | Matrix & chart test | **IN_ZIP** |
| `monthly_reporting.py` | SOP classification & DB cloning | Bundled in ZIP | `standalone/monthly_reporting.py` | Unit test | **IN_ZIP** |
| `pipeline_common.py` | Site extraction & safe saving | Bundled in ZIP | `standalone/pipeline_common.py` | Unit test | **IN_ZIP** |
| `Output/write_db_month.py` | Stage 3 database population | Bundled in ZIP | `standalone/Output/write_db_month.py` | DB sheet sync test | **IN_ZIP** |
| `Output/update_copq_pptx.py` | Stage 4 executive slide update | Bundled in ZIP | `standalone/Output/update_copq_pptx.py` | PPTX update test | **IN_ZIP** |
| `agents/` modules | Pipeline orchestrator & sentinel | Bundled in ZIP | `standalone/agents/` | Orchestration test | **IN_ZIP** |
| `Color_Defect.xlsx` | Defect color fills for charts | Bundled in ZIP | `standalone/BC_Color/Color_Defect.xlsx` | Embedded color check | **IN_ZIP** |
| `CoPQ database 25.xlsx` | Master cost database | Bundled in ZIP | `standalone/Database/CoPQ database 25.xlsx` | Overview sheet check | **IN_ZIP** |
| `CoPQ_type_analysis.xlsx` | Master type analysis database | Bundled in ZIP | `standalone/Database/CoPQ_type_analysis.xlsx` | Cost ranking check | **IN_ZIP** |
| `Template_COPQ.pptx` | Base slide presentation | Bundled in ZIP | `standalone/Template_COPQ/Template_COPQ.pptx` | Slide deck load check | **IN_ZIP** |
| Synthetic Fixtures | Non-destructive verification | Bundled in ZIP | `fixtures/` | End-to-end execution | **IN_ZIP** |

### 24.2 Handoff Coverage Table
| Pipeline Stage | Transfer Status | Reproduction Method |
|---|---|---|
| Stage 1: COPQ Ingestion & Cleaning | **REPRODUCIBLE_FROM_ARTIFACT** | Fully runnable from `standalone/combine_files.py` and `fixtures/COPQ_Input/` |
| Stage 2: FTT Defect Processing & Charts | **REPRODUCIBLE_FROM_ARTIFACT** | Fully runnable from `standalone/process_ftt.py`, `Color_Defect.xlsx`, and `fixtures/FTT_Input/` |
| Stage 3: Database Synchronization | **REPRODUCIBLE_FROM_ARTIFACT** | Fully runnable from `standalone/Output/write_db_month.py` and `Database/*.xlsx` |
| Stage 4: PowerPoint Presentation Update | **REPRODUCIBLE_FROM_ARTIFACT** | Fully runnable from `standalone/Output/update_copq_pptx.py` and `Template_COPQ.pptx` |
| Data Quality Sentinel & Telemetry | **REPRODUCIBLE_FROM_ARTIFACT** | Fully runnable from `standalone/agents/` |

---

## 25. Reference fixture and parity assertions

### 25.1 Synthetic Fixture Description
The `fixtures/` folder contains sanitized, minimal test data designed to verify all business rules without confidential data:
- `fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-VH.xlsx`: Contains 4 data rows testing B/C defect, Touch-up remark override, Rework defect, and direct Touch-up.
- `fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-JV.xlsx`: Contains 2 data rows testing B/C defect and Re-inspection.
- `fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-VH.xlsx`: Contains 4 rows testing Grade B/C filtering and `Model 2` priority.
- `fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-JV.xlsx`: Contains 3 rows testing Grade B/C filtering and fallback ranking.

### 25.2 Parity Assertion Commands
To verify baseline parity against the synthetic fixtures:
```bash
python standalone/combine_files.py --input fixtures/COPQ_Input --output test_out/COPQ_Clean_test.xlsx
python standalone/process_ftt.py --ftt-dir fixtures/FTT_Input --output test_out/FTT_Combined_test.xlsx --color-map standalone/BC_Color/Color_Defect.xlsx --cost-file standalone/Database/CoPQ_type_analysis.xlsx
```
**Expected Parity Invariants**:
1. `COPQ_Clean_test.xlsx` has sheet `COPQ_Clean` with exactly 5 combined data rows.
2. In `COPQ_Clean_test.xlsx`, row with remark "Touch-up paint required" has `Type` = `"Touch-up"`.
3. `FTT_Combined_test.xlsx` contains sheet `Top Defects Summary` with stacked BarChart and formula cells matching `=IFERROR(SUMIFS(`.

---

## 26. Source/asset capsules and bundle manifest

### 26.1 Bundle Manifest Table
| Path in Package | Role / Description | Size (Bytes) | SHA-256 Checksum | Original Relative Path | Requirement Level |
|---|---|---|---|---|---|
| `standalone/combine_files.py` | Core pipeline 1: COPQ file cleaner & consolidator | 24991 | `d44e7655d16b8a4e609bebba1bef44e0475281eade47c3a87b94338c75de0c03` | `combine_files.py` | Required at Runtime |
| `standalone/process_ftt.py` | Core pipeline 2: FTT defect extractor & chart builder | 34777 | `c291765991174696c5557e937bbd4ef5ec9c6527712a46bf0b9a4ff4ebe76a7c` | `process_ftt.py` | Required at Runtime |
| `standalone/monthly_reporting.py` | Core SOP classification & database cloning helper | 7224 | `8738e35c6ba7582df98b7123dc8007d64cdf6caa5f1e3f92ab0fed996c6f131c` | `monthly_reporting.py` | Required at Runtime |
| `standalone/pipeline_common.py` | Shared utilities: site extraction & safe file saving | 2456 | `5d372c478852d16f62ce4eeaef30d944af52bd1efad79b3c4fbbff6d5ffc093f` | `pipeline_common.py` | Required at Runtime |
| `standalone/Output/Combined_COPQ.py` | COPQ Clean consolidation & formatting helper | 13662 | `9700b5d2e282d1425e5530d8f4721e05f4dea36031ce655557d6efbc207e963b` | `Output/Combined_COPQ.py` | Required at Runtime |
| `standalone/Output/COPQ_Type_Detail.py` | COPQ Type Detail sheet builder | 5031 | `d79899ee0ae6c850807f9aafb4b5ee290c17a247aca51d638c0f8f53517841e0` | `Output/COPQ_Type_Detail.py` | Required at Runtime |
| `standalone/Output/COPQ_Pivot_table.py` | COPQ Pivot Table overview builder | 21078 | `4fab8715649b4c755a7d57c71c943061ade8d2e9fc961b52731c624077db022c` | `Output/COPQ_Pivot_table.py` | Required at Runtime |
| `standalone/Output/COPQ_Database.py` | COPQ Database report category builder | 13207 | `69368876945b856eba5923847c8bcc356b725435ac1e743ba4c16f808a2b0e13` | `Output/COPQ_Database.py` | Required at Runtime |
| `standalone/Output/write_db_month.py` | Core pipeline 3: Monthly database sheet population script | 32117 | `c3ed5158d5ef1ebcb596aeada3a0e5162f5a6f011ae028439c36c88adbd4dcec` | `Output/write_db_month.py` | Required at Runtime |
| `standalone/Output/update_copq_pptx.py` | Core pipeline 4: Executive PowerPoint deck update engine | 78255 | `75849a84e1e4c00ca6923162fb60f9384158d406e5f9f1cb617dcff4c4027630` | `Output/update_copq_pptx.py` | Required at Runtime |
| `standalone/agents/__init__.py` | Autonomous agents package initializer | 708 | `fc7cf9b3fd38ff51766c46ef57f94cc642bee90d8c2b1c532141cbf011f584e0` | `agents/__init__.py` | Required at Runtime |
| `standalone/agents/pipeline_orchestrator.py` | Atomic 4-stage pipeline orchestrator agent | 12998 | `faae9ee3d742acb08337fb142aef4811404e9a270e059e36d6ebdeb9526c4339` | `agents/pipeline_orchestrator.py` | Required at Runtime |
| `standalone/agents/quality_sentinel.py` | Data quality sentinel & defect dictionary auditor agent | 15615 | `eb263d65896bea81a89295a148adc4057ad733db3d39186a440546e323b8fcd9` | `agents/quality_sentinel.py` | Required at Runtime |
| `standalone/agents/watchdog_delivery.py` | Executive briefing generator & directory watcher daemon | 16147 | `68357bb82f9cf917004813cdaf8a3459e0698c7ac94557fd41a67d0229a062fd` | `agents/watchdog_delivery.py` | Required at Runtime |
| `standalone/agents/run_agent.py` | Unified CLI runner for all autonomous agent workflows | 4279 | `62199ca20ca808b241572d584c52b14273b4a18417a775fe74ae484aeab99b15` | `agents/run_agent.py` | Required at Runtime |
| `standalone/BC_Color/Color_Defect.xlsx` | Static reference asset: Defect-to-color mapping workbook | 15330 | `b70212a17b44ef811067a8244e063b5f85a3772223eaa60a94553c545f01793b` | `BC_Color/Color_Defect.xlsx` | Required at Runtime |
| `standalone/Database/CoPQ database 25.xlsx` | Master database: Overview costs & site monthly trend workbook | 65291 | `15d62e1a38a497e51537db210acbd2bb6fe6d4766ff38bfabf68383466159e7a` | `Database/CoPQ database 25.xlsx` | Required at Runtime |
| `standalone/Database/CoPQ_type_analysis.xlsx` | Master database: Type analysis & model cost ranking workbook | 229402 | `f8fc155c53ac2567bdd630be4aa8b6416593b7671e89363ebc6e9783d689a7a4` | `Database/CoPQ_type_analysis.xlsx` | Required at Runtime |
| `standalone/Template_COPQ/Template_COPQ.pptx` | Base presentation template: Executive PowerPoint deck | 663020 | `8a04793dcb1f1b5f24836243e154d8c278d2b5e2701e76548992225af8072282` | `Template_COPQ/Template_COPQ.pptx` | Required at Runtime |
| `fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-VH.xlsx` | Synthetic sanitized COPQ ERP export for VH | 6203 | `b602761a09445794a03ab8abec83fc72a4809224219f0b322393fb4fe22935f3` | `scratch/test_fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-VH.xlsx` | Verification Only |
| `fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-JV.xlsx` | Synthetic sanitized COPQ ERP export for JV | 5896 | `6b36de03396d4a84d9413f4037fd9d493caa4436425e3551fdf0f4eccad57d8f` | `scratch/test_fixtures/COPQ_Input/Cost Of Poor Quality  2026-08-01_2026-08-31-JV.xlsx` | Verification Only |
| `fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-VH.xlsx` | Synthetic sanitized FTT defect export for VH | 5922 | `8d58d721b518626f76ed61c8e29e09f229f0fa28e424935594e9ee63314cf33e` | `scratch/test_fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-VH.xlsx` | Verification Only |
| `fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-JV.xlsx` | Synthetic sanitized FTT defect export for JV | 5804 | `6d7474d5eafc9bad03e6e03e7dccfe760d6d039f689f697813e38e23969cc8e1` | `scratch/test_fixtures/FTT_Input/FTT-Internal_2026-08-01_2026-08-31-JV.xlsx` | Verification Only |
| `fixtures/expected_summary.json` | Machine-readable expected parity invariants | 878 | `876570f40cf11a13d9fba2280804fa2addc71556637e3b2bf13b88d17167a771` | `generated` | Verification Only |

---

## 27. Receiver rehearsal and archive verification

### 27.1 Rehearsal Verification Procedure
To prove cross-person portability, the extractor executed a receiver rehearsal:
1. Created a fresh, isolated temporary directory.
2. Extracted `copq_pipeline_MCP_HANDOFF.zip` into the sandbox.
3. Computed SHA-256 hashes of all extracted files and confirmed 100% bitwise parity with Section 26.1.
4. Executed `combine_files.py` and `process_ftt.py` against `fixtures/` within the isolated directory.
5. Confirmed that all output workbooks were generated, verified formulas, and confirmed zero references to the host machine's original filesystem.
6. Cleaned up the temporary sandbox.

---

## 28. Definition of done for AI #2

An implementer integrating this project into the target MCP server repository has completed the task when:
1. **Tool Registration**: The MCP tool `copq_pipeline` is registered with the schemas defined in Section 17.
2. **Algorithmic Parity**:
   - Touch-up remark override rule is preserved.
   - Dynamic Excel `=SUMIFS` formulas and openpyxl BarCharts are generated with correct colors.
   - Dynamic period detection works for `YYYY-MM` formats.
3. **In-Process Execution**: All modules are invoked in-process without spawning desktop GUI processes or shell scripts.
4. **Locked-File Resilience**: `save_with_fallback()` is preserved so open files do not crash the service.
5. **Automated Parity Test**: A test suite verifies that processing the provided synthetic fixtures reproduces the invariants in Section 25.
