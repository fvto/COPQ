#!/usr/bin/env python3
"""Compatibility shim for the canonical Output implementation."""

import os
import runpy

if __name__ == "__main__":
    script_path = os.path.join(os.path.dirname(__file__), "Output", "update_copq_pptx.py")
    if not os.path.exists(script_path):
        raise FileNotFoundError(f"Canonical script not found: {script_path}")
    runpy.run_path(script_path, run_name="__main__")
    raise SystemExit(0)
