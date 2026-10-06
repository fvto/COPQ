# COPQ — Cost Of Poor Quality Toolkit

A Windows-based toolset for combining, processing, and visualizing **COPQ (Cost Of Poor Quality)** data exported from the ERP system, plus an **FTT B/C report** generator and an **interactive dashboard** with live filters and charts.

---

## 📁 Project Structure

```
combine_COPQ/
├── Interactive_Dashboard.py      # Main app: interactive Tkinter dashboard
├── Run_Interactive_Dashboard.bat # One-click launcher (checks env & installs deps)
├── combine_files.py              # COPQ combiner & 4-step workflow orchestrator
├── process_ftt.py                # FTT B/C report processor
├── monthly_reporting.py          # Type classification helpers
├── pipeline_common.py            # Common utilities (site name extraction, safe save)
├── Source/                       # Core pipeline modules & reporting engines
│   ├── Combined_COPQ.py          # Step 1: Script COPQ Clean
│   ├── COPQ_Type_Detail.py       # Step 2: Script COPQ Type Detail
│   ├── COPQ_Pivot_table.py       # Step 3: Script COPQ Pivot Table
│   ├── COPQ_Database.py          # Step 4: COPQ Database
│   ├── write_db_month.py         # Monthly database consolidation engine
│   ├── update_copq_pptx.py       # Executive PowerPoint update engine
│   ├── monthly_reporting.py      # Classification & validation helpers
│   └── generate_pptx.py          # Legacy slide presentation builder
├── COPQ_Input/                   # ⬅ Drop raw ERP COPQ exports here (.xls/.xlsx/.csv)
├── FTT_Input/                    # ⬅ Drop raw FTT reports here
├── BC_Color/Color_Defect.xlsx    # Color/defect mapping used by the FTT processor
├── Database/                     # Monthly database reference workbooks
└── Output/                       # ⬅ Pure generated reports & deliverables
    ├── COPQ_Clean.xlsx           # Combined master workbook (12 sheets)
    ├── FTT_Combined_Report.xlsx  # Combined FTT B/C report
    └── COPQ_Report_*.pptx        # Monthly executive PowerPoint presentations
```

---

## 🚀 Quick Start

1. Double-click **`Run_Interactive_Dashboard.bat`**.
   - It verifies Python 3.9+ is in PATH.
   - Auto-installs missing packages: `pandas`, `openpyxl`, `numpy`, `matplotlib`.
   - Launches the interactive dashboard.

2. Or run manually:
   ```powershell
   python Interactive_Dashboard.py
   ```

### Command-line pipeline (optional)

```powershell
python combine_files.py     # combine COPQ_Input/* → Output/COPQ_Clean.xlsx
python process_ftt.py       # FTT_Input/* → Output/FTT_Combined_Report.xlsx
```

---

## 🖥️ Interactive Dashboard

### Data loading
- Loads `Output/COPQ_Clean.xlsx` automatically on startup if it exists.
- **Browse** supports selecting **multiple files at once** — raw ERP exports are cleaned and column-mapped on the fly (same logic as the combine pipeline), with the site auto-detected from each filename (`VH`, `VH2`, `VH3`, `VH4`, `JV`, `JV2`, `JV3`, `JVB`). Supported formats: `.xlsx`, `.xls`, `.csv`.

### Filters
| Filter | Options |
|---|---|
| Month | `All` or a specific `YYYY-MM` |
| Metric | `Cost ($)` · `Qty (pairs)` · `All` (shows both $ and pairs) |
| Defect type | Touch-up, Rework, Reinspection, B/C |
| Site | VH, VH2, VH3, VH4, JV, JV2, JV3, JVB (+ All/None toggle) |

All KPI cards, charts, and the data table update live when any filter changes.

### KPI cards
TOTAL FILTERED · TOUCH-UP · REWORK · REINSPECTION · B/C GRADE · RECORDS

With Metric = **All**, each card shows both `$cost | qty pairs`.

### Charts (Charts tab)
1. **Monthly Trend by Type** — stacked bars per month with total labels
2. **Type Share** — donut chart of defect-type distribution
3. **Site Ranking** — horizontal bar ranking by site
4. **Top Models** — top 8 models by selected metric

### Filtered Data tab
A scrollable table of the filtered records (first 1000 rows).

### Automation panel
| Button | Action |
|---|---|
| **1. Combine COPQ** | Combines all files in `COPQ_Input/` → `Output/COPQ_Clean.xlsx` |
| **2. Process FTT** | Builds `Output/FTT_Combined_Report.xlsx` from `FTT_Input/` |
| **Run Full Pipeline** | Runs steps 1 + 2, then refreshes the dashboard |
| **Open Output Folder** | Opens `Output/` in Explorer |

Each run shows a **progress dialog** with step counter, progress bar, and percentage.

---

## 📊 Generated Output Files

- `Output/COPQ_Clean.xlsx` — combined master with sheets:
  - Per-site cleaned raw data (JV, JV2, JV3, VH, VH2, VH3, VH4)
  - `COPQ_Clean` — master table with `Site` column prepended
  - `COPQ_Pivot_Overview` — aggregated per site & type
  - `COPQ_Pivot_TopFactory` — aggregated per site, factory & type
  - `COPQ_Pivot_TopModel` — B/C, reinspection & touch-up by site & model
  - `COPQ_Report` — executive cost summary & breakdown sub-tables (Touch-up / Re-inspection / B/C / Rework)
- `Output/FTT_Combined_Report.xlsx` — FTT B/C report with color mapping

---

## 🧠 How classification works

Records are classified into **Touch-up / Rework / Reinspection / B/C** using keywords found in `Type`, `Type Detail`, and `Remark (Re-inspection Station)` columns (e.g., "touch" → Touch-up, "reinsp" → Reinspection, "rework" → Rework, "b/c"/"grade" → B/C).

---

## ✅ Requirements

- Windows (uses `os.startfile`)
- Python 3.9+
- Packages: `pandas`, `numpy`, `openpyxl`, `matplotlib`, `python-pptx`, `mcp`
  ```powershell
  pip install pandas numpy openpyxl matplotlib python-pptx mcp
  ```

---

## 🤖 Model Context Protocol (MCP) Server

The toolkit exposes a production-grade, standard MCP server (`mcp_server.py`) enabling AI coding assistants, autonomous agents, and IDEs (e.g. Claude Desktop, Antigravity IDE, Cursor) to execute the end-to-end pipeline in-process with structured outputs.

### Launching the MCP Server
```powershell
python mcp_server.py                  # Standard stdio transport
python mcp_server.py --transport sse  # SSE transport on port 8000
```

### Registered Tools
| Tool | Action / Description | Key Parameters |
|---|---|---|
| `copq_pipeline` | Primary orchestrator tool | `action` (`run_all`, `combine_copq`, `process_ftt`, `update_database`, `generate_presentation`, `audit_quality`), `period`, `copq_input_dir`, `ftt_input_dir`, `output_dir` |
| `copq_audit_quality` | Pre-flight quality sentinel audit across all 7 manufacturing plants | `period`, `copq_input_dir`, `ftt_input_dir` |
| `copq_generate_briefing` | Generates C-suite briefing (Markdown & styled HTML) | `period`, `output_dir` |

### Registered Resources
- `copq://specs/implementation_spec` — Authoritative COPQ MCP Implementation Specification
- `copq://rules/pipeline_rules` — SOP Defect Classification & Multi-Site Normalization Rules
- `copq://reports/quality_audit` — Latest Data Quality Sentinel Audit Report
- `copq://reports/executive_briefing` — Latest C-Suite Executive Briefing
- `copq://reports/telemetry` — Pipeline execution telemetry and runtime metrics

### Client Configuration (`mcp_config.json` snippet)
```json
{
  "mcpServers": {
    "copq_pipeline": {
      "command": "python",
      "args": ["C:\\Users\\User\\Desktop\\combine_COPQ\\mcp_server.py"]
    }
  }
}
```

### Running Automated Parity Tests
```powershell
python -m pytest tests/test_copq_mcp.py -v
```
All 8 verification tests (schema contracts, touch-up override rule, site precedence, synthetic fixture end-to-end parity, and negative error handling) execute without error.

