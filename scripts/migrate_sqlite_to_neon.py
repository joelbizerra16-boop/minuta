#!/usr/bin/env python
"""
Fase N3 — Migracao controlada SQLite -> PostgreSQL (Neon).

Fluxo ETL:
  Inventario -> Extracao (read-only) -> Validacao -> Transformacao -> Load -> Validacao -> Relatorio

Uso:
  set MINUTA_SQLITE_SOURCE_URL=sqlite:///C:/caminho/minuta.db
  set MINUTA_DATABASE_URL=postgresql+psycopg2://...?sslmode=require

  python scripts/migrate_sqlite_to_neon.py --dry-run
  python scripts/migrate_sqlite_to_neon.py --execute
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.migration.report import MigrationReport
from scripts.migration.runner import run_migration

REPORT_PATH = PROJECT_ROOT / "reports" / "migration_n3_report.json"


def main() -> int:
    parser = argparse.ArgumentParser(description="Migracao SQLite -> PostgreSQL (Neon)")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--dry-run", action="store_true", help="Inventario + validacao sem carga")
    group.add_argument("--execute", action="store_true", help="Executa migracao completa")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)-5s [%(name)s] %(message)s")

    report: MigrationReport = run_migration(dry_run=args.dry_run)
    saved = report.save(REPORT_PATH)

    if report.aprovada:
        logging.getLogger("minuta.migration").info(
            "FASE N3 %s — relatorio em %s",
            "DRY-RUN OK" if args.dry_run else "CONCLUIDA",
            saved,
        )
        return 0

    logging.getLogger("minuta.migration").error(
        "FASE N3 FALHOU — bloqueador=%s relatorio=%s",
        report.bloqueador,
        saved,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
