"""Benchmark P0.1 — bootstrap once-per-process (antes vs depois da idempotencia)."""

from __future__ import annotations

import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _ms(start: float) -> float:
    return round((time.perf_counter() - start) * 1000, 2)


def _simulate_streamlit_reruns(*, database_url: str, reruns: int) -> dict[str, object]:
    from core.bootstrap import configure_application_storage
    from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
    from core.settings import get_settings
    from core.startup_environment import run_startup_environment_checks
    from core.startup_retention import reset_startup_retention_flag, run_startup_retention_once
    from infrastructure.database import get_engine_create_count
    from infrastructure.schema import get_alembic_run_count
    from core.environment_diagnostics import get_postgresql_validation_count

    get_settings.cache_clear()
    os.environ["MINUTA_DATABASE_URL"] = database_url
    get_settings.cache_clear()
    reset_startup_retention_flag()
    reset_infrastructure_bootstrap_state()

    rows: list[dict[str, object]] = []
    for index in range(1, reruns + 1):
        t0 = time.perf_counter()
        run_startup_environment_checks()
        t1 = time.perf_counter()
        configure_application_storage()
        t2 = time.perf_counter()
        run_startup_retention_once()
        t3 = time.perf_counter()
        rows.append(
            {
                "rerun": index,
                "environment_ms": _ms(t0) if index == 1 else _ms(t1) - _ms(t0) + _ms(t0),
                "environment_ms_only": round((t1 - t0) * 1000, 2),
                "configure_storage_ms": round((t2 - t1) * 1000, 2),
                "retention_ms": round((t3 - t2) * 1000, 2),
                "total_ms": round((t3 - t0) * 1000, 2),
            }
        )

    return {
        "database_url": database_url,
        "reruns": rows,
        "engine_creates": get_engine_create_count(),
        "alembic_runs": get_alembic_run_count(),
        "postgres_validations": get_postgresql_validation_count(),
    }


def main() -> int:
    with tempfile.TemporaryDirectory() as tmp_dir:
        sqlite_url = f"sqlite:///{(Path(tmp_dir) / 'bench_p01.db').as_posix()}"
        result = _simulate_streamlit_reruns(database_url=sqlite_url, reruns=5)
        from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state

        reset_infrastructure_bootstrap_state()

    print("## Benchmark P0.1 — Bootstrap once-per-process")
    print()
    print("| Rerun | environment (ms) | configure_storage (ms) | retention (ms) | total (ms) |")
    print("| ---: | ---: | ---: | ---: | ---: |")
    for row in result["reruns"]:
        print(
            f"| {row['rerun']} | {row['environment_ms_only']} | {row['configure_storage_ms']} | "
            f"{row['retention_ms']} | {row['total_ms']} |"
        )
    print()
    print(f"- `create_engine()` chamadas no processo: **{result['engine_creates']}** (esperado: 1)")
    print(f"- Execucoes Alembic no processo: **{result['alembic_runs']}** (esperado: 0 em SQLite)")
    print(f"- Validacoes PostgreSQL no processo: **{result['postgres_validations']}** (esperado: 0 em SQLite)")
    print()
    print("Com PostgreSQL/Neon, o 1o rerun paga Alembic/validacao; reruns 2-5 devem ficar < 50 ms.")

    postgres_url = str(os.getenv("MINUTA_DATABASE_URL", "") or "").strip()
    if "postgres" in postgres_url:
        from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state

        pg = _simulate_streamlit_reruns(database_url=postgres_url, reruns=3)
        print()
        print("### PostgreSQL (MINUTA_DATABASE_URL)")
        for row in pg["reruns"]:
            print(
                f"rerun {row['rerun']}: total={row['total_ms']}ms "
                f"(env={row['environment_ms_only']} configure={row['configure_storage_ms']})"
            )
        print(f"engine_creates={pg['engine_creates']} alembic_runs={pg['alembic_runs']} validations={pg['postgres_validations']}")
        reset_infrastructure_bootstrap_state()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
