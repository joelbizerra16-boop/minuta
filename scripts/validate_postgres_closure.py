#!/usr/bin/env python
"""Validacao rapida PostgreSQL para encerramento N4."""
from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from sqlalchemy import inspect, text

from core.bootstrap import configure_application_storage
from infrastructure.database import get_engine


def main() -> int:
    configure_application_storage()
    engine = get_engine()
    assert engine.dialect.name == "postgresql", f"Esperado postgresql, recebido {engine.dialect.name}"

    with engine.connect() as conn:
        version = str(conn.scalar(text("SELECT version()")) or "")[:80]
        database = str(conn.scalar(text("SELECT current_database()")) or "")
        revision = str(conn.scalar(text("SELECT version_num FROM alembic_version LIMIT 1")) or "")
        fk_count = int(
            conn.scalar(
                text(
                    "SELECT COUNT(*) FROM information_schema.table_constraints "
                    "WHERE constraint_type='FOREIGN KEY' AND table_schema='public'"
                )
            )
            or 0
        )
        idx_count = int(conn.scalar(text("SELECT COUNT(*) FROM pg_indexes WHERE schemaname='public'")) or 0)
        tables = [t for t in inspect(engine).get_table_names() if t != "alembic_version"]
        counts = {}
        total = 0
        for table in tables:
            n = int(conn.scalar(text(f'SELECT COUNT(*) FROM "{table}"')) or 0)
            counts[table] = n
            total += n

    print(f"motor=postgresql database={database}")
    print(f"version={version}")
    print(f"alembic_head={revision}")
    print(f"tables={len(tables)} records={total} fk={fk_count} indexes={idx_count}")
    for name in sorted(counts):
        print(f"  {name}={counts[name]}")
  # documento_xml check
    xml_fs = sum(1 for _ in (Path("C:/MinutaData/xml_storage").rglob("*.xml")))
    print(f"xml_filesystem={xml_fs}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
