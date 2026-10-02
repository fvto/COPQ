"""
COPQ Proactive Agent Automation Suite
=====================================
Workflows and autonomous agents solving the 3 core inefficiencies:
1. Pipeline Orchestrator Agent (Inefficiency 1: Fragmented, brittle multi-step pipeline)
2. Quality Sentinel Agent (Inefficiency 2: Passive quality checks & silent defect omission)
3. Watchdog & Executive Delivery Agent (Inefficiency 3: Reactive manual triggers & delayed intelligence)
"""

from .pipeline_orchestrator import PipelineOrchestratorAgent
from .quality_sentinel import QualitySentinelAgent
from .watchdog_delivery import WatchdogDeliveryAgent

__all__ = [
    "PipelineOrchestratorAgent",
    "QualitySentinelAgent",
    "WatchdogDeliveryAgent",
]
