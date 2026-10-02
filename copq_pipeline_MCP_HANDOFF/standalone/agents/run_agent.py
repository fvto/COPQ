#!/usr/bin/env python3
"""
COPQ Autonomous Agent CLI Runner
=================================
Unified CLI entry point for the 3 Proactive Agents & Workflows:
1. orchestrate: Run atomic 4-stage pipeline (COPQ Combine, FTT Process, DB Update, PPTX Generation).
2. audit: Run proactive data quality & defect dictionary audit.
3. brief: Generate C-suite executive intelligence briefing (Markdown & HTML).
4. watch: Start proactive input file watcher / daemon.
5. run-all: Complete autonomous cycle (Audit -> Orchestrate -> Brief).

Usage:
  python agents/run_agent.py orchestrate [--period YYYY-MM]
  python agents/run_agent.py audit [--period YYYY-MM]
  python agents/run_agent.py brief [--period YYYY-MM]
  python agents/run_agent.py watch [--poll-interval 10] [--once]
  python agents/run_agent.py run-all [--period YYYY-MM]
"""

import os
import sys
import argparse

# Ensure root path is in sys.path
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from agents.pipeline_orchestrator import PipelineOrchestratorAgent
from agents.quality_sentinel import QualitySentinelAgent
from agents.watchdog_delivery import WatchdogDeliveryAgent


def main():
    parser = argparse.ArgumentParser(
        description="COPQ Proactive Agent Automation Runner",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    subparsers = parser.add_subparsers(dest="command", help="Agent workflow to invoke")

    # Command: orchestrate
    p_orch = subparsers.add_parser("orchestrate", help="Run atomic end-to-end pipeline")
    p_orch.add_argument("--period", default=None, help="Reporting period (e.g. 2026-08 or 'Aug, 2026')")
    p_orch.add_argument("--copq-dir", default=None, help="Custom COPQ input directory")
    p_orch.add_argument("--ftt-dir", default=None, help="Custom FTT input directory")
    p_orch.add_argument("--output-dir", default=None, help="Custom Output directory")

    # Command: audit
    p_audit = subparsers.add_parser("audit", help="Run proactive data quality & anomaly sentinel audit")
    p_audit.add_argument("--period", default=None, help="Reporting period label")

    # Command: brief
    p_brief = subparsers.add_parser("brief", help="Generate executive intelligence briefing")
    p_brief.add_argument("--period", default="Current Period", help="Reporting period label")

    # Command: watch
    p_watch = subparsers.add_parser("watch", help="Start continuous directory watcher daemon")
    p_watch.add_argument("--poll-interval", type=int, default=10, help="Polling interval in seconds")
    p_watch.add_argument("--once", action="store_true", help="Run a single check & deliver cycle, then exit")

    # Command: run-all
    p_all = subparsers.add_parser("run-all", help="Execute complete autonomous cycle (Audit -> Orchestrate -> Delivery)")
    p_all.add_argument("--period", default=None, help="Reporting period (e.g. 2026-08)")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    if args.command == "orchestrate":
        agent = PipelineOrchestratorAgent(ROOT_DIR)
        res = agent.execute_pipeline(
            period_override=args.period,
            copq_dir=args.copq_dir,
            ftt_dir=args.ftt_dir,
            output_dir=args.output_dir
        )
        sys.exit(0 if res.get("status") == "SUCCESS" else 1)

    elif args.command == "audit":
        agent = QualitySentinelAgent(ROOT_DIR)
        res = agent.run_full_audit(period_label=args.period)
        sys.exit(0 if res.get("status") != "ACTION_REQUIRED" else 1)

    elif args.command == "brief":
        agent = WatchdogDeliveryAgent(ROOT_DIR)
        res = agent.generate_executive_briefing(period_label=args.period)
        sys.exit(0 if res else 1)

    elif args.command == "watch":
        agent = WatchdogDeliveryAgent(ROOT_DIR)
        agent.watch_inputs(poll_interval_seconds=args.poll_interval, run_once=args.once)
        sys.exit(0)

    elif args.command == "run-all":
        agent = WatchdogDeliveryAgent(ROOT_DIR)
        res = agent.process_and_deliver(period_override=args.period)
        sys.exit(0 if res.get("telemetry", {}).get("status") == "SUCCESS" else 1)


if __name__ == "__main__":
    main()
