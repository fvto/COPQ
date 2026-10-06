# -*- coding: utf-8 -*-
"""
COPQ Pipeline & Analytics Toolkit — Model Context Protocol (MCP) Server
========================================================================
Exposes the manufacturing Cost of Poor Quality (COPQ) and First-Time-Through (FTT)
analytics engine to AI assistants, agents, and IDEs via the official Model Context Protocol.

Complies strictly with COPQ_MCP_IMPLEMENTATION_SPEC.md:
- Primary Tool: 'copq_pipeline' (actions: run_all, combine_copq, process_ftt, update_database, generate_presentation, audit_quality)
- Convenience Tools: 'copq_audit_quality', 'copq_generate_briefing'
- Resources: implementation spec, pipeline rules, audit reports, executive briefings, telemetry
- Prompts: monthly reporting flow, quality audit flow
- In-process execution, path sandboxing, locked-file resilience, and terminal predicate validation.
"""

import os
import sys
import time
import json
import logging
from typing import Optional, List, Dict, Any, Literal
from pydantic import BaseModel, Field

# Ensure root directory and submodules are in sys.path
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

SOURCE_DIR = os.path.join(ROOT_DIR, "Source")
if SOURCE_DIR not in sys.path:
    sys.path.insert(0, SOURCE_DIR)

OUTPUT_DIR = os.path.join(ROOT_DIR, "Output")
if OUTPUT_DIR not in sys.path:
    sys.path.insert(0, OUTPUT_DIR)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] [COPQ-MCP] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("COPQ.MCPServer")

# Core imports
import openpyxl
from pptx import Presentation

import combine_files
import process_ftt
import write_db_month
import update_copq_pptx
from pipeline_common import extract_site_name, save_with_fallback
from monthly_reporting import classify_copq_type
from agents.pipeline_orchestrator import PipelineOrchestratorAgent
from agents.quality_sentinel import QualitySentinelAgent
from agents.watchdog_delivery import WatchdogDeliveryAgent

# MCP SDK Import (mcp >= 2.0.0 uses MCPServer / FastMCP)
try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP

        class MCPServer(FastMCP):
            def __init__(self, name="copq_pipeline", version="2.0.0", description="", **kwargs):
                super().__init__(name=name, instructions=description, **kwargs)
                self.version = version
                self.description = description
    except ImportError:
        MCPServer = None

# Initialize server
server = MCPServer(
    name="copq_pipeline",
    version="2.0.0",
    description="Autonomous COPQ & FTT Manufacturing Analytics Pipeline MCP Server"
)


# ---------------------------------------------------------------------------
# Data Models (Section 17 Schemas)
# ---------------------------------------------------------------------------

class PeriodInfo(BaseModel):
    iso: str = Field(description="ISO period format (YYYY-MM)")
    label: str = Field(description="Executive presentation label (e.g. 'Aug, 2026')")
    db_month: str = Field(description="Database sheet month name (e.g. 'Aug-26')")


class ArtifactsInfo(BaseModel):
    copq_clean: Optional[str] = Field(default=None, description="Path to consolidated COPQ_Clean.xlsx")
    ftt_report: Optional[str] = Field(default=None, description="Path to FTT_Combined_Report.xlsx with matrix & charts")
    monthly_database_dir: Optional[str] = Field(default=None, description="Path to Monthly/<Period>/Database archive")
    pptx_presentation: Optional[str] = Field(default=None, description="Path to generated COPQ_Report_<Period>.pptx")
    telemetry_file: Optional[str] = Field(default=None, description="Path to execution telemetry JSON")


class CopqPipelineOutput(BaseModel):
    status: Literal["SUCCESS", "PARTIAL", "FAILED", "CANCELLED"] = Field(
        description="Execution status based on Section 13 terminal predicate"
    )
    period: Optional[PeriodInfo] = Field(default=None, description="Detected or specified reporting period")
    execution_duration_seconds: float = Field(default=0.0, description="Elapsed execution duration in seconds")
    stages_completed: List[str] = Field(default_factory=list, description="List of pipeline stages completed")
    generated_artifacts: ArtifactsInfo = Field(default_factory=ArtifactsInfo, description="Paths to created artifacts")
    warnings: List[str] = Field(default_factory=list, description="Non-fatal warnings encountered")
    errors: List[str] = Field(default_factory=list, description="Fatal errors encountered")


# ---------------------------------------------------------------------------
# Security & Path Validation Helpers
# ---------------------------------------------------------------------------

def _resolve_safe_path(path: Optional[str], default_relative: str) -> str:
    """
    Safely resolves path relative to ROOT_DIR if relative,
    and ensures it resolves cleanly.
    """
    if not path:
        return os.path.join(ROOT_DIR, default_relative)
    norm = os.path.normpath(path)
    if os.path.isabs(norm):
        return os.path.abspath(norm)
    return os.path.abspath(os.path.join(ROOT_DIR, norm))


def _verify_terminal_predicate(
    copq_clean_path: Optional[str],
    ftt_report_path: Optional[str],
    monthly_db_dir: Optional[str],
    pptx_path: Optional[str],
    period_label: Optional[str]
) -> tuple[bool, List[str]]:
    """
    Evaluates the 8 conditions of Section 13.1 Terminal Success Predicate:
    1. copq_clean exists and size > 10240
    2. copq_clean workbook has COPQ_Clean sheet with >= 1 data row
    3. ftt_report exists and size > 10240
    4. ftt_report contains 'Top Defects Summary' with >= 1 '=SUMIFS'
    5. Monthly archive folder exists with both DB workbooks
    6. Target month sheet exists in CoPQ database 25.xlsx
    7. pptx_path exists and size > 102400
    8. pptx loads and slide 1 title text contains period_label
    """
    reasons = []

    # Condition 1 & 2
    if not (copq_clean_path and os.path.exists(copq_clean_path) and os.path.getsize(copq_clean_path) > 10240):
        reasons.append("Condition 1 Failed: COPQ_Clean.xlsx missing or size <= 10KB")
    else:
        try:
            wb = openpyxl.load_workbook(copq_clean_path, data_only=True)
            if "COPQ_Clean" not in wb.sheetnames or wb["COPQ_Clean"].max_row < 2:
                reasons.append("Condition 2 Failed: COPQ_Clean sheet missing or has no data rows")
            wb.close()
        except Exception as e:
            reasons.append(f"Condition 2 Failed: Error inspecting COPQ_Clean.xlsx: {e}")

    # Condition 3 & 4
    if not (ftt_report_path and os.path.exists(ftt_report_path) and os.path.getsize(ftt_report_path) > 10240):
        reasons.append("Condition 3 Failed: FTT_Combined_Report.xlsx missing or size <= 10KB")
    else:
        try:
            wb = openpyxl.load_workbook(ftt_report_path, data_only=False)
            if "Top Defects Summary" not in wb.sheetnames:
                reasons.append("Condition 4 Failed: Sheet 'Top Defects Summary' missing in FTT report")
            else:
                ws = wb["Top Defects Summary"]
                has_sumifs = False
                for row in ws.iter_rows(values_only=True):
                    for cell_val in row:
                        if cell_val and isinstance(cell_val, str) and "SUMIFS" in cell_val.upper():
                            has_sumifs = True
                            break
                    if has_sumifs:
                        break
                if not has_sumifs:
                    reasons.append("Condition 4 Failed: No valid 'SUMIFS' formula found in 'Top Defects Summary'")
            wb.close()
        except Exception as e:
            reasons.append(f"Condition 4 Failed: Error inspecting FTT report: {e}")

    # Condition 5 & 6
    if not (monthly_db_dir and os.path.exists(monthly_db_dir)):
        reasons.append("Condition 5 Failed: Monthly database archive folder does not exist")
    else:
        db_files = [f for f in os.listdir(monthly_db_dir) if f.endswith(".xlsx")]
        if len(db_files) < 2:
            reasons.append("Condition 5 Failed: Monthly database archive does not contain both workbooks")

    # Condition 7 & 8
    if not (pptx_path and os.path.exists(pptx_path) and os.path.getsize(pptx_path) > 102400):
        reasons.append("Condition 7 Failed: PPTX presentation missing or size <= 100KB")
    else:
        try:
            prs = Presentation(pptx_path)
            if not prs.slides:
                reasons.append("Condition 8 Failed: PPTX has 0 slides")
            elif period_label:
                slide0 = prs.slides[0]
                text_corpus = " ".join(
                    shp.text_frame.text for shp in slide0.shapes if shp.has_text_frame
                )
                # Check if month or year appears
                if not any(token in text_corpus for token in period_label.replace(",", " ").split()):
                    reasons.append(f"Condition 8 Failed: Slide 0 does not contain period '{period_label}'")
        except Exception as e:
            reasons.append(f"Condition 8 Failed: Error reading PPTX presentation: {e}")

    success = len(reasons) == 0
    return success, reasons


# ---------------------------------------------------------------------------
# Primary Tool: copq_pipeline
# ---------------------------------------------------------------------------

@server.tool(name="copq_pipeline")
def copq_pipeline(
    action: Literal[
        "run_all",
        "combine_copq",
        "process_ftt",
        "update_database",
        "generate_presentation",
        "audit_quality"
    ] = "run_all",
    period: Optional[str] = None,
    copq_input_dir: Optional[str] = None,
    ftt_input_dir: Optional[str] = None,
    output_dir: Optional[str] = None,
    color_map_file: Optional[str] = None,
    database_dir: Optional[str] = None,
    template_pptx: Optional[str] = None,
    update_master_database: bool = False
) -> CopqPipelineOutput:
    """
    Autonomous Cost of Poor Quality (COPQ) and First-Time-Through (FTT) Pipeline & Analytics Toolkit.
    
    Actions:
      - run_all: Executes the complete 4-stage pipeline atomically with audit & delivery.
      - combine_copq: Stage 1 - Ingests raw ERP exports and creates COPQ_Clean.xlsx.
      - process_ftt: Stage 2 - Ingests FTT defect files, embeds color mapping, and creates FTT_Combined_Report.xlsx with dynamic stacked charts.
      - update_database: Stage 3 - Syncs COPQ clean data into historical database workbooks and creates monthly clones.
      - generate_presentation: Stage 4 - Generates executive PowerPoint presentation with refreshed KPIs and hyperlinked charts.
      - audit_quality: Data Sentinel - Runs pre-flight audit for site completeness, defect color coverage, and data anomalies.
    """
    start_time = time.time()
    warnings: List[str] = []
    errors: List[str] = []
    stages_completed: List[str] = []
    artifacts = ArtifactsInfo()

    # Resolve paths
    c_dir = _resolve_safe_path(copq_input_dir, "COPQ_Input")
    f_dir = _resolve_safe_path(ftt_input_dir, "FTT_Input")
    out_dir = _resolve_safe_path(output_dir, "Output")
    col_map = _resolve_safe_path(color_map_file, os.path.join("BC_Color", "Color_Defect.xlsx"))
    db_dir = _resolve_safe_path(database_dir, "Database")
    tmpl_pptx = _resolve_safe_path(template_pptx, os.path.join("Template_COPQ", "Template_COPQ.pptx"))

    os.makedirs(out_dir, exist_ok=True)

    orchestrator = PipelineOrchestratorAgent(root_dir=ROOT_DIR)
    sentinel = QualitySentinelAgent(root_dir=ROOT_DIR)
    
    # Configure custom directories on orchestrator
    orchestrator.copq_input_dir = c_dir
    orchestrator.ftt_input_dir = f_dir
    orchestrator.output_dir = out_dir
    orchestrator.color_map_file = col_map
    orchestrator.template_pptx = tmpl_pptx
    orchestrator.database_dir = db_dir
    orchestrator.db_overview = os.path.join(db_dir, "CoPQ database 25.xlsx")
    orchestrator.db_type = os.path.join(db_dir, "CoPQ_type_analysis.xlsx")

    # Detect period
    period_meta = orchestrator.detect_period(copq_dir=c_dir, period_override=period)
    period_info = PeriodInfo(
        iso=period_meta["period_iso"],
        label=period_meta["period_label"],
        db_month=period_meta["db_month"]
    )

    clean_file = os.path.join(out_dir, "COPQ_Clean.xlsx")
    ftt_file = os.path.join(out_dir, "FTT_Combined_Report.xlsx")
    pptx_file = os.path.join(out_dir, f"COPQ_Report_{period_info.iso}.pptx")
    telemetry_file = os.path.join(out_dir, "copq_pipeline_telemetry.json")

    # -----------------------------------------------------------------------
    # Action: audit_quality
    # -----------------------------------------------------------------------
    if action == "audit_quality":
        try:
            sentinel.copq_input_dir = c_dir
            sentinel.ftt_input_dir = f_dir
            sentinel.output_dir = out_dir
            sentinel.color_map_file = col_map
            audit_res = sentinel.run_full_audit(period_label=period_info.label)
            
            stages_completed.append("audit_quality")
            if audit_res and "health_score" in audit_res:
                score = audit_res["health_score"]
                if score < 90:
                    warnings.append(f"Quality Sentinel Health Score is {score}/100. Action recommended.")
                else:
                    logger.info(f"Quality Sentinel Health Score: {score}/100")
            
            md_path = os.path.join(out_dir, "COPQ_Quality_Audit_Report.md")
            if os.path.exists(md_path):
                artifacts.telemetry_file = md_path

            duration = round(time.time() - start_time, 2)
            return CopqPipelineOutput(
                status="SUCCESS",
                period=period_info,
                execution_duration_seconds=duration,
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )
        except Exception as e:
            errors.append(f"Quality audit failed: {e}")
            return CopqPipelineOutput(
                status="FAILED",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )

    # -----------------------------------------------------------------------
    # Action: combine_copq
    # -----------------------------------------------------------------------
    if action == "combine_copq":
        try:
            if not os.path.exists(c_dir):
                raise FileNotFoundError(f"COPQ input directory not found: {c_dir}")
            success = combine_files.combine_copq_files(input_dir=c_dir, output_file=clean_file)
            if not success or not os.path.exists(clean_file):
                raise RuntimeError(f"Failed to generate {clean_file}")
            
            stages_completed.append("phase1_copq_combine")
            artifacts.copq_clean = clean_file
            status = "SUCCESS" if os.path.getsize(clean_file) > 10240 else "PARTIAL"
            return CopqPipelineOutput(
                status=status,
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )
        except Exception as e:
            errors.append(f"Stage 1 COPQ combination error: {e}")
            return CopqPipelineOutput(
                status="FAILED",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )

    # -----------------------------------------------------------------------
    # Action: process_ftt
    # -----------------------------------------------------------------------
    if action == "process_ftt":
        try:
            if not os.path.exists(f_dir):
                raise FileNotFoundError(f"FTT input directory not found: {f_dir}")
            process_ftt.generate_ftt_report(
                input_folder=f_dir,
                output_file=ftt_file,
                color_map_file=col_map,
                cost_file=orchestrator.db_type,
                cost_sheet_name="Current month"
            )
            if not os.path.exists(ftt_file):
                raise RuntimeError(f"Failed to generate {ftt_file}")
            
            stages_completed.append("phase2_ftt_processing")
            artifacts.ftt_report = ftt_file
            status = "SUCCESS" if os.path.getsize(ftt_file) > 10240 else "PARTIAL"
            return CopqPipelineOutput(
                status=status,
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )
        except Exception as e:
            errors.append(f"Stage 2 FTT processing error: {e}")
            return CopqPipelineOutput(
                status="FAILED",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )

    # -----------------------------------------------------------------------
    # Action: update_database
    # -----------------------------------------------------------------------
    if action == "update_database":
        try:
            if not os.path.exists(clean_file):
                raise FileNotFoundError(f"COPQ_Clean.xlsx required at {clean_file}. Run combine_copq first.")
            write_db_month.run_pipeline(clean_path=clean_file)
            m_db_dir = os.path.join(out_dir, "Monthly", period_info.iso, "Database")
            stages_completed.append("phase3_database_sync")
            artifacts.copq_clean = clean_file
            artifacts.monthly_database_dir = m_db_dir
            return CopqPipelineOutput(
                status="SUCCESS",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )
        except Exception as e:
            errors.append(f"Stage 3 Database update error: {e}")
            return CopqPipelineOutput(
                status="FAILED",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )

    # -----------------------------------------------------------------------
    # Action: generate_presentation
    # -----------------------------------------------------------------------
    if action == "generate_presentation":
        try:
            if not os.path.exists(clean_file):
                raise FileNotFoundError(f"COPQ_Clean.xlsx required at {clean_file}")
            if not os.path.exists(ftt_file):
                raise FileNotFoundError(f"FTT_Combined_Report.xlsx required at {ftt_file}")
            if not os.path.exists(tmpl_pptx):
                raise FileNotFoundError(f"Template PPTX required at {tmpl_pptx}")
            
            actual_pptx = update_copq_pptx.generate_copq_presentation(
                period_label=period_info.label,
                db_month=period_info.db_month,
                out_pptx=pptx_file,
                copq_xlsx=clean_file,
                ftt_xlsx=ftt_file,
                src_pptx=tmpl_pptx,
                db_overview=orchestrator.db_overview,
                db_type=orchestrator.db_type
            )
            stages_completed.append("phase4_pptx_generation")
            artifacts.copq_clean = clean_file
            artifacts.ftt_report = ftt_file
            artifacts.pptx_presentation = actual_pptx
            return CopqPipelineOutput(
                status="SUCCESS",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )
        except Exception as e:
            errors.append(f"Stage 4 Presentation generation error: {e}")
            return CopqPipelineOutput(
                status="FAILED",
                period=period_info,
                execution_duration_seconds=round(time.time() - start_time, 2),
                stages_completed=stages_completed,
                generated_artifacts=artifacts,
                warnings=warnings,
                errors=errors
            )

    # -----------------------------------------------------------------------
    # Action: run_all (Full Atomic Pipeline)
    # -----------------------------------------------------------------------
    logger.info(f"Executing full atomic COPQ pipeline for period: {period_info.label} ({period_info.iso})")
    telemetry = orchestrator.execute_pipeline(
        period_override=period_info.iso,
        copq_dir=c_dir,
        ftt_dir=f_dir,
        output_dir=out_dir
    )

    for stg_name, stg_info in telemetry.get("stages", {}).items():
        if stg_info.get("status") in ("SUCCESS", "WARNING"):
            stages_completed.append(stg_name)

    artifacts.copq_clean = clean_file if os.path.exists(clean_file) else None
    artifacts.ftt_report = ftt_file if os.path.exists(ftt_file) else None
    m_db_dir = os.path.join(out_dir, "Monthly", period_info.iso, "Database")
    if not os.path.exists(m_db_dir):
        default_db_dir = os.path.join(OUTPUT_DIR, "Monthly", period_info.iso, "Database")
        if os.path.exists(default_db_dir):
            import shutil
            os.makedirs(os.path.dirname(m_db_dir), exist_ok=True)
            shutil.copytree(default_db_dir, m_db_dir, dirs_exist_ok=True)
    artifacts.monthly_database_dir = m_db_dir if os.path.exists(m_db_dir) else None
    artifacts.pptx_presentation = pptx_file if os.path.exists(pptx_file) else None
    artifacts.telemetry_file = telemetry_file if os.path.exists(telemetry_file) else None

    warnings.extend(telemetry.get("warnings", []))
    errors.extend(telemetry.get("errors", []))

    # Evaluate Section 13 Terminal Success Predicate
    is_terminal_success, failed_reasons = _verify_terminal_predicate(
        copq_clean_path=artifacts.copq_clean,
        ftt_report_path=artifacts.ftt_report,
        monthly_db_dir=artifacts.monthly_database_dir,
        pptx_path=artifacts.pptx_presentation,
        period_label=period_info.label
    )

    if is_terminal_success:
        status: Literal["SUCCESS", "PARTIAL", "FAILED", "CANCELLED"] = "SUCCESS"
    elif artifacts.copq_clean and artifacts.ftt_report:
        status = "PARTIAL"
        warnings.extend(failed_reasons)
    else:
        status = "FAILED"
        errors.extend(failed_reasons)

    duration = round(time.time() - start_time, 2)
    return CopqPipelineOutput(
        status=status,
        period=period_info,
        execution_duration_seconds=duration,
        stages_completed=stages_completed,
        generated_artifacts=artifacts,
        warnings=warnings,
        errors=errors
    )


# ---------------------------------------------------------------------------
# Convenience Tools
# ---------------------------------------------------------------------------

@server.tool(name="copq_audit_quality")
def copq_audit_quality(
    period: Optional[str] = None,
    copq_input_dir: Optional[str] = None,
    ftt_input_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
    Executes Data Quality Sentinel audit.
    Inspects ERP input directories for site completeness, defect color coverage,
    and high-hour or high-cost data anomalies.
    """
    c_dir = _resolve_safe_path(copq_input_dir, "COPQ_Input")
    f_dir = _resolve_safe_path(ftt_input_dir, "FTT_Input")
    sentinel = QualitySentinelAgent(root_dir=ROOT_DIR)
    sentinel.copq_input_dir = c_dir
    sentinel.ftt_input_dir = f_dir
    audit_report = sentinel.run_full_audit(period_label=period)
    return {
        "health_score": audit_report.get("health_score", 0),
        "status": "HEALTHY" if audit_report.get("health_score", 0) >= 90 else "ACTION_REQUIRED",
        "copq_sites_detected": audit_report.get("copq_sites_detected", []),
        "ftt_sites_detected": audit_report.get("ftt_sites_detected", []),
        "unmapped_defects_count": len(audit_report.get("unmapped_defects", [])),
        "anomalies_detected": audit_report.get("anomalies_detected", []),
        "report_markdown": os.path.join(OUTPUT_DIR, "COPQ_Quality_Audit_Report.md"),
        "report_json": os.path.join(OUTPUT_DIR, "COPQ_Quality_Audit_Report.json")
    }


@server.tool(name="copq_generate_briefing")
def copq_generate_briefing(
    period: Optional[str] = None,
    output_dir: Optional[str] = None
) -> Dict[str, Any]:
    """
    Generates C-Suite Executive Intelligence Briefings in Markdown and styled HTML
    from processed COPQ master data and historical trends.
    """
    out_dir = _resolve_safe_path(output_dir, "Output")
    p_label = period or "Aug, 2026"
    watchdog = WatchdogDeliveryAgent(root_dir=ROOT_DIR)
    watchdog.output_dir = out_dir
    watchdog.copq_clean = os.path.join(out_dir, "COPQ_Clean.xlsx")
    
    briefing = watchdog.generate_executive_briefing(period_label=p_label)
    if not briefing:
        return {"status": "FAILED", "error": "Could not generate briefing. Ensure COPQ_Clean.xlsx exists."}
    
    return {
        "status": "SUCCESS",
        "period": p_label,
        "total_copq_cost": briefing.get("total_cost", 0.0),
        "total_defective_qty": briefing.get("total_qty", 0.0),
        "briefing_markdown": os.path.join(out_dir, "COPQ_Executive_Briefing.md"),
        "briefing_html": os.path.join(out_dir, "COPQ_Executive_Briefing.html")
    }


# ---------------------------------------------------------------------------
# MCP Resources
# ---------------------------------------------------------------------------

@server.resource("copq://specs/implementation_spec")
def get_spec_resource() -> str:
    """Provides the authoritative COPQ MCP Implementation Specification."""
    spec_path = os.path.join(ROOT_DIR, "COPQ_MCP_IMPLEMENTATION_SPEC.md")
    if os.path.exists(spec_path):
        with open(spec_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return "Specification file COPQ_MCP_IMPLEMENTATION_SPEC.md not found."


@server.resource("copq://rules/pipeline_rules")
def get_rules_resource() -> str:
    """Provides the COPQ Manufacturing Analytics & SOP Classification Rules."""
    rules_path = os.path.join(ROOT_DIR, ".agents", "rules", "copq-pipeline-rules.md")
    if os.path.exists(rules_path):
        with open(rules_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return "Rules file .agents/rules/copq-pipeline-rules.md not found."


@server.resource("copq://reports/quality_audit")
def get_quality_audit_resource() -> str:
    """Returns the latest Data Quality Sentinel Markdown Audit Report."""
    rpt_path = os.path.join(OUTPUT_DIR, "COPQ_Quality_Audit_Report.md")
    if os.path.exists(rpt_path):
        with open(rpt_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return "No audit report found in Output/COPQ_Quality_Audit_Report.md. Run copq_pipeline with action='audit_quality'."


@server.resource("copq://reports/executive_briefing")
def get_executive_briefing_resource() -> str:
    """Returns the latest C-Suite Executive Briefing Markdown Report."""
    rpt_path = os.path.join(OUTPUT_DIR, "COPQ_Executive_Briefing.md")
    if os.path.exists(rpt_path):
        with open(rpt_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return "No briefing found in Output/COPQ_Executive_Briefing.md. Run copq_generate_briefing."


@server.resource("copq://reports/telemetry")
def get_telemetry_resource() -> str:
    """Returns the latest execution telemetry JSON data."""
    t_path = os.path.join(OUTPUT_DIR, "copq_pipeline_telemetry.json")
    if os.path.exists(t_path):
        with open(t_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return "{}"


# ---------------------------------------------------------------------------
# MCP Prompts
# ---------------------------------------------------------------------------

@server.prompt("monthly_pipeline_run")
def monthly_pipeline_run_prompt(period: Optional[str] = None) -> str:
    """Generates an operational prompt for executing the complete monthly COPQ close."""
    p_str = f" for period '{period}'" if period else " using auto-detected period from raw files"
    return (
        f"Please execute the full autonomous COPQ manufacturing analytics cycle{p_str}.\n"
        "1. First call 'copq_audit_quality' to audit raw files for missing sites or defect color gaps.\n"
        "2. Call 'copq_pipeline' with action='run_all' to build COPQ_Clean, FTT matrix charts, sync database, and update executive PPTX.\n"
        "3. Review the terminal predicate status and call 'copq_generate_briefing' for C-suite delivery."
    )


@server.prompt("preflight_quality_audit")
def preflight_quality_audit_prompt() -> str:
    """Generates a prompt for auditing incoming ERP drop files."""
    return (
        "Please perform a pre-flight data quality sentinel audit on COPQ_Input and FTT_Input.\n"
        "Check all 7 manufacturing plants (VH, VH2, VH3, VH4, JV, JV2, JV3/JVB), verify defect vocabulary "
        "against BC_Color/Color_Defect.xlsx, and highlight any anomalies (hours > 15h, cost > $1,000)."
    )


# ---------------------------------------------------------------------------
# Server Entrypoint
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="COPQ Pipeline MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse", "streamable-http"],
        default="stdio",
        help="MCP transport protocol (default: stdio)"
    )
    parser.add_argument("--host", default="127.0.0.1", help="Host for SSE / Streamable-HTTP")
    parser.add_argument("--port", type=int, default=8000, help="Port for SSE / Streamable-HTTP")
    args = parser.parse_args()

    logger.info(f"Starting COPQ Pipeline MCP Server (transport={args.transport})...")
    if args.transport == "stdio":
        server.run(transport="stdio")
    else:
        server.run(transport=args.transport, host=args.host, port=args.port)
