#!/usr/bin/env python
"""
Benchmark N1/N2: compara tempos SQLite (local) vs PostgreSQL/Neon (MINUTA_DATABASE_URL).

Uso:
  python scripts/neon_benchmark_compare.py
  set MINUTA_DATABASE_URL=postgresql+psycopg2://...  (opcional para comparar Neon)
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

from sqlalchemy import text

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _reset_engine() -> None:
    import infrastructure.database as db_module

    if db_module._engine is not None:
        db_module._engine.dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def _benchmark_label(database_url: str, label: str) -> dict[str, float]:
    from core.bootstrap import configure_application_storage
    from core.settings import get_settings
    from infrastructure.database import get_engine

    get_settings.cache_clear()
    os.environ["MINUTA_DATABASE_URL"] = database_url
    get_settings.cache_clear()

    with tempfile.TemporaryDirectory() as tmp_dir:
        data_root = Path(tmp_dir)
        os.environ["MINUTA_DATABASE_URL"] = database_url
        get_settings.cache_clear()

        bootstrap_start = time.perf_counter()
        configure_application_storage()
        timings: dict[str, float] = {
            "bootstrap_ms": round((time.perf_counter() - bootstrap_start) * 1000, 2),
        }
        engine = get_engine()
        connect_start = time.perf_counter()
        with engine.connect() as connection:
            timings["conexao_ms"] = round((time.perf_counter() - connect_start) * 1000, 2)

            for query_label, sql in {
                "select_1": "SELECT 1",
                "count_usuario": "SELECT COUNT(*) FROM usuario",
            }.items():
                start = time.perf_counter()
                try:
                    connection.scalar(text(sql))
                except Exception:
                    timings[query_label] = -1.0
                else:
                    timings[query_label] = round((time.perf_counter() - start) * 1000, 2)

        _reset_engine()
        get_settings.cache_clear()
        timings["_label"] = label  # type: ignore[assignment]
        return timings


def main() -> int:
    sqlite_url = f"sqlite:///{(Path(tempfile.gettempdir()) / 'minuta_bench.db').as_posix()}"
    results = [_benchmark_label(sqlite_url, "sqlite")]

    postgres_url = str(os.getenv("MINUTA_DATABASE_URL", "") or "").strip()
    if postgres_url and "postgres" in postgres_url:
        results.append(_benchmark_label(postgres_url, "postgresql"))
    else:
        print("MINUTA_DATABASE_URL PostgreSQL nao definida — apenas SQLite medido.")

    print("## Benchmark N1/N2 (ms)")
    print()
    print("| Motor | Bootstrap | Conexao | SELECT 1 | COUNT usuario |")
    print("| --- | ---: | ---: | ---: | ---: |")
    for row in results:
        label = row.pop("_label")
        print(
            f"| {label} | {row.get('bootstrap_ms', 'n/a')} | {row.get('conexao_ms', 'n/a')} | "
            f"{row.get('select_1', 'n/a')} | {row.get('count_usuario', 'n/a')} |"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
