#!/usr/bin/env python
"""Diagnostico do ambiente (FASE 0) — sem migracao nem alteracao de banco."""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.environment_diagnostics import run_environment_diagnostics


def main() -> int:
    report = run_environment_diagnostics()
    print(report.format_banner())
    return 0 if report.ready_for_migration or "OPERACIONAL" in report.overall_status else 1


if __name__ == "__main__":
    raise SystemExit(main())
