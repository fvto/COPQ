"""
Pipeline Orchestrator Agent (Workflow 1)
=========================================
Solves Inefficiency 1:
- Eliminates fragmented, brittle multi-step script executions and manual UI handoffs.
- Auto-detects reporting period dynamically (eliminating hardcoded month strings).
- Coordinates end-to-end execution atomically across COPQ Combining, FTT Analysis,
  Database Synchronization, and Executive PPTX Deck Generation.
- Handles locked files gracefully and exports machine-readable telemetry.
"""

import os
import sys
import re
import time
import json
import logging
from datetime import datetime, date

# Configure logger
logger = logging.getLogger("COPQ.PipelineOrchestrator")
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] [Orchestrator] %(message)s", "%H:%M:%S"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

# Root directory resolution
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

SOURCE_DIR = os.path.join(ROOT_DIR, "Source")
if SOURCE_DIR not in sys.path:
    sys.path.insert(0, SOURCE_DIR)

OUTPUT_DIR_DEFAULT = os.path.join(ROOT_DIR, "Output")
if OUTPUT_DIR_DEFAULT not in sys.path:
    sys.path.insert(0, OUTPUT_DIR_DEFAULT)

# Import underlying workflow engines
import combine_files
import process_ftt
import write_db_month
import update_copq_pptx


class PipelineOrchestratorAgent:
    """Proactive Agent managing end-to-end COPQ and FTT pipeline execution."""

    MONTH_MAP = {
        1: ("Jan", "January"), 2: ("Feb", "February"), 3: ("Mar", "March"),
        4: ("Apr", "April"), 5: ("May", "May"), 6: ("Jun", "June"),
        7: ("Jul", "July"), 8: ("Aug", "August"), 9: ("Sep", "September"),
        10: ("Oct", "October"), 11: ("Nov", "November"), 12: ("Dec", "December")
    }

    def __init__(self, root_dir=ROOT_DIR):
        self.root_dir = root_dir
        self.copq_input_dir = os.path.join(root_dir, "COPQ_Input")
        self.ftt_input_dir = os.path.join(root_dir, "FTT_Input")
        self.output_dir = os.path.join(root_dir, "Output")
        self.database_dir = os.path.join(root_dir, "Database")
        self.color_map_file = os.path.join(root_dir, "BC_Color", "Color_Defect.xlsx")
        self.template_pptx = os.path.join(root_dir, "Template_COPQ", "Template_COPQ.pptx")
        self.db_overview = os.path.join(self.database_dir, "CoPQ database 25.xlsx")
        self.db_type = os.path.join(self.database_dir, "CoPQ_type_analysis.xlsx")

    def detect_period(self, copq_dir=None, period_override=None):
        """
        Dynamically detects the reporting period from filenames or input data.
        Returns a dictionary with period metadata:
          - year (int)
          - month (int)
          - period_iso: 'YYYY-MM'
          - period_label: 'Aug, 2026'
          - db_month: 'Aug-2026'
          - month_name: 'August'
        """
        if period_override:
            # Parse 'YYYY-MM' or 'Month, YYYY'
            m_iso = re.match(r"^(\d{4})-(\d{1,2})$", period_override.strip())
            if m_iso:
                year, month = int(m_iso.group(1)), int(m_iso.group(2))
                short_m, full_m = self.MONTH_MAP[month]
                return {
                    "year": year,
                    "month": month,
                    "period_iso": f"{year:04d}-{month:02d}",
                    "period_label": f"{short_m}, {year}",
                    "db_month": f"{short_m}-{year}",
                    "month_name": full_m,
                }

        search_dir = copq_dir or self.copq_input_dir
        candidates = []
        if os.path.exists(search_dir):
            for fname in os.listdir(search_dir):
                # Search pattern like 2026-08-01 or 2026_08
                m = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", fname)
                if m:
                    candidates.append((int(m.group(1)), int(m.group(2))))
                else:
                    m2 = re.search(r"(\d{4})[-_](\d{2})", fname)
                    if m2:
                        candidates.append((int(m2.group(1)), int(m2.group(2))))

        if candidates:
            # Pick most frequent or max date
            year, month = max(candidates)
            short_m, full_m = self.MONTH_MAP[month]
            return {
                "year": year,
                "month": month,
                "period_iso": f"{year:04d}-{month:02d}",
                "period_label": f"{short_m}, {year}",
                "db_month": f"{short_m}-{year}",
                "month_name": full_m,
            }

        # Fallback to current year/month
        now = datetime.now()
        short_m, full_m = self.MONTH_MAP[now.month]
        return {
            "year": now.year,
            "month": now.month,
            "period_iso": f"{now.year:04d}-{now.month:02d}",
            "period_label": f"{short_m}, {now.year}",
            "db_month": f"{short_m}-{now.year}",
            "month_name": full_m,
        }

    def execute_pipeline(self, period_override=None, copq_dir=None, ftt_dir=None, output_dir=None):
        """
        Executes the entire 4-phase COPQ and FTT processing pipeline atomically.
        """
        start_time = time.time()
        c_dir = copq_dir or self.copq_input_dir
        f_dir = ftt_dir or self.ftt_input_dir
        out_dir = output_dir or self.output_dir
        os.makedirs(out_dir, exist_ok=True)

        period = self.detect_period(copq_dir=c_dir, period_override=period_override)
        logger.info(f"Target Reporting Period: {period['period_label']} ({period['period_iso']})")

        copq_clean_xlsx = os.path.join(out_dir, "COPQ_Clean.xlsx")
        ftt_combined_xlsx = os.path.join(out_dir, "FTT_Combined_Report.xlsx")
        pptx_output = os.path.join(out_dir, f"COPQ_Report_{period['period_iso']}.pptx")

        telemetry = {
            "orchestrator_version": "2.0.0",
            "execution_timestamp": datetime.now().isoformat(),
            "period": period,
            "stages": {},
            "status": "RUNNING",
            "errors": [],
            "warnings": [],
            "generated_artifacts": {}
        }

        # ---------------------------------------------------------
        # Phase 1: Combine & Clean COPQ Data
        # ---------------------------------------------------------
        t0 = time.time()
        logger.info(">>> Phase 1/4: Combining & Cleaning COPQ Data...")
        try:
            success = combine_files.combine_copq_files(
                input_dir=c_dir,
                output_file=copq_clean_xlsx
            )
            if not success or not os.path.exists(copq_clean_xlsx):
                raise RuntimeError(f"COPQ combination failed or output file not created: {copq_clean_xlsx}")
            
            telemetry["stages"]["phase1_copq_combine"] = {
                "status": "SUCCESS",
                "duration_seconds": round(time.time() - t0, 2),
                "output_file": copq_clean_xlsx,
                "size_bytes": os.path.getsize(copq_clean_xlsx)
            }
            telemetry["generated_artifacts"]["copq_clean"] = copq_clean_xlsx
            logger.info(f"Phase 1 Complete: Saved {copq_clean_xlsx} ({round(time.time() - t0, 2)}s)")
        except Exception as e:
            msg = f"Phase 1 Error: {str(e)}"
            logger.error(msg)
            telemetry["stages"]["phase1_copq_combine"] = {"status": "FAILED", "error": msg}
            telemetry["status"] = "FAILED"
            telemetry["errors"].append(msg)
            self._save_telemetry(telemetry, out_dir)
            return telemetry

        # ---------------------------------------------------------
        # Phase 2: Process FTT B/C Defect Data & Charts
        # ---------------------------------------------------------
        t0 = time.time()
        logger.info(">>> Phase 2/4: Processing FTT B/C Defect Matrix & Openpyxl Charts...")
        try:
            # Cost file points to CoPQ_type_analysis
            process_ftt.generate_ftt_report(
                input_folder=f_dir,
                output_file=ftt_combined_xlsx,
                color_map_file=self.color_map_file,
                cost_file=self.db_type,
                cost_sheet_name="Current month"
            )
            if not os.path.exists(ftt_combined_xlsx):
                raise RuntimeError(f"FTT processing failed or output not created: {ftt_combined_xlsx}")

            telemetry["stages"]["phase2_ftt_processing"] = {
                "status": "SUCCESS",
                "duration_seconds": round(time.time() - t0, 2),
                "output_file": ftt_combined_xlsx,
                "size_bytes": os.path.getsize(ftt_combined_xlsx)
            }
            telemetry["generated_artifacts"]["ftt_report"] = ftt_combined_xlsx
            logger.info(f"Phase 2 Complete: Saved {ftt_combined_xlsx} ({round(time.time() - t0, 2)}s)")
        except Exception as e:
            msg = f"Phase 2 Error: {str(e)}"
            logger.error(msg)
            telemetry["stages"]["phase2_ftt_processing"] = {"status": "FAILED", "error": msg}
            telemetry["status"] = "FAILED"
            telemetry["errors"].append(msg)
            self._save_telemetry(telemetry, out_dir)
            return telemetry

        # ---------------------------------------------------------
        # Phase 3: Synchronize Historical Database Workbooks
        # ---------------------------------------------------------
        t0 = time.time()
        logger.info(">>> Phase 3/4: Synchronizing Master Database Workbooks...")
        try:
            # write_db_month runs pipeline with clean_path
            write_db_month.run_pipeline(clean_path=copq_clean_xlsx)
            monthly_archive = os.path.join(out_dir, "Monthly", period["period_iso"], "Database")
            telemetry["stages"]["phase3_database_sync"] = {
                "status": "SUCCESS",
                "duration_seconds": round(time.time() - t0, 2),
                "database_updated": [self.db_overview, self.db_type],
                "monthly_archive_dir": monthly_archive
            }
            telemetry["generated_artifacts"]["database_clones"] = monthly_archive
            logger.info(f"Phase 3 Complete: Databases updated and cloned to {monthly_archive} ({round(time.time() - t0, 2)}s)")
        except Exception as e:
            msg = f"Phase 3 Warning (Non-Fatal): {str(e)}"
            logger.warning(msg)
            telemetry["stages"]["phase3_database_sync"] = {"status": "WARNING", "error": msg}
            telemetry["warnings"].append(msg)

        # ---------------------------------------------------------
        # Phase 4: Executive PowerPoint Slide Deck Generation
        # ---------------------------------------------------------
        t0 = time.time()
        logger.info(f">>> Phase 4/4: Generating Executive Presentation ({period['period_label']})...")
        try:
            actual_pptx = update_copq_pptx.generate_copq_presentation(
                period_label=period["period_label"],
                db_month=period["db_month"],
                out_pptx=pptx_output,
                copq_xlsx=copq_clean_xlsx,
                ftt_xlsx=ftt_combined_xlsx,
                src_pptx=self.template_pptx,
                db_overview=self.db_overview,
                db_type=self.db_type
            )
            telemetry["stages"]["phase4_pptx_generation"] = {
                "status": "SUCCESS",
                "duration_seconds": round(time.time() - t0, 2),
                "output_file": actual_pptx,
                "size_bytes": os.path.getsize(actual_pptx) if os.path.exists(actual_pptx) else 0
            }
            telemetry["generated_artifacts"]["pptx_report"] = actual_pptx
            logger.info(f"Phase 4 Complete: Presentation written to {actual_pptx} ({round(time.time() - t0, 2)}s)")
        except Exception as e:
            msg = f"Phase 4 Error: {str(e)}"
            logger.error(msg)
            telemetry["stages"]["phase4_pptx_generation"] = {"status": "FAILED", "error": msg}
            telemetry["errors"].append(msg)

        total_duration = round(time.time() - start_time, 2)
        telemetry["total_duration_seconds"] = total_duration
        telemetry["status"] = "SUCCESS" if not telemetry["errors"] else "COMPLETED_WITH_ERRORS"
        logger.info(f"=== PIPELINE RUN COMPLETE in {total_duration}s (Status: {telemetry['status']}) ===")

        self._save_telemetry(telemetry, out_dir)
        return telemetry

    def _save_telemetry(self, telemetry, out_dir):
        t_path = os.path.join(out_dir, "copq_pipeline_telemetry.json")
        try:
            with open(t_path, "w", encoding="utf-8") as f:
                json.dump(telemetry, f, indent=2)
            logger.info(f"Telemetry saved: {t_path}")
        except Exception as e:
            logger.warning(f"Failed to write telemetry: {e}")
