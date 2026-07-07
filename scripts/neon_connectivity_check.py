#!/usr/bin/env python
"""
Homologacao N1/N2: conectividade Neon PostgreSQL sem migracao de dados.

Uso:
  set MINUTA_DATABASE_URL=postgresql+psycopg2://...
  python scripts/neon_connectivity_check.py

Requer variavel de ambiente MINUTA_DATABASE_URL apontando para o Neon.
"""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path

from sqlalchemy import inspect, text
from sqlalchemy.engine import make_url

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)-5s [%(name)s] %(message)s",
)
_LOGGER = logging.getLogger("minuta.neon_check")

_EXPECTED_TABLES = frozenset(
    {
        "perfil",
        "usuario",
        "motorista",
        "veiculo",
        "destinatario",
        "rota",
        "nota_fiscal",
        "item_nota_fiscal",
        "carregamento",
        "item_carregamento",
        "documento",
        "historico_operacional",
        "evento_auditoria",
        "configuracao",
        "documento_xml",
        "alembic_version",
    }
)


def _require_postgres_url() -> str:
    database_url = str(os.getenv("MINUTA_DATABASE_URL", "") or "").strip()
    if not database_url:
        raise SystemExit(
            "MINUTA_DATABASE_URL nao definida. "
            "Configure postgresql+psycopg2://usuario:senha@host/database?sslmode=require"
        )
    driver = make_url(database_url).drivername or ""
    if "postgresql" not in driver and "postgres" not in driver:
        raise SystemExit(f"MINUTA_DATABASE_URL deve apontar para PostgreSQL/Neon (recebido: {driver})")
    return database_url


def _log_url_summary(database_url: str) -> None:
    url = make_url(database_url)
    query = dict(url.query) if url.query else {}
    sslmode = query.get("sslmode", "nao_informado")
    _LOGGER.info(
        "config database=%s host=%s port=%s sslmode=%s",
        url.database,
        url.host,
        url.port or 5432,
        sslmode,
    )


def _reset_engine() -> None:
    import infrastructure.database as db_module

    if db_module._engine is not None:
        db_module._engine.dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def _test_connection(engine) -> dict[str, float | str]:
    timings: dict[str, float | str] = {}
    start = time.perf_counter()
    with engine.connect() as connection:
        timings["conexao_ms"] = round((time.perf_counter() - start) * 1000, 2)
        version = connection.scalar(text("SELECT version()"))
        timings["postgresql_version"] = str(version or "")[:120]
        current_db = connection.scalar(text("SELECT current_database()"))
        timings["current_database"] = str(current_db or "")
        ssl_row = connection.execute(
            text("SELECT ssl FROM pg_stat_ssl WHERE pid = pg_backend_pid()")
        ).first()
        timings["ssl_ativo"] = "sim" if ssl_row and ssl_row[0] else "nao"
    return timings


def _test_read_write_rollback(engine) -> None:
    with engine.begin() as connection:
        probe = connection.execute(
            text("SELECT 1 AS ok, current_user AS usuario")
        ).mappings().one()
        assert int(probe["ok"]) == 1
        _LOGGER.info("leitura ok usuario=%s", probe["usuario"])

    with engine.begin() as connection:
        connection.execute(
            text(
                "CREATE TABLE IF NOT EXISTS minuta_neon_probe ("
                "id SERIAL PRIMARY KEY, payload TEXT NOT NULL)"
            )
        )
        connection.execute(
            text("INSERT INTO minuta_neon_probe (payload) VALUES (:payload)"),
            {"payload": "neon-homologacao"},
        )
        count = connection.scalar(text("SELECT COUNT(*) FROM minuta_neon_probe"))
        _LOGGER.info("escrita ok registros_probe=%s", count)

    with engine.connect() as connection:
        trans = connection.begin()
        try:
            connection.execute(
                text("INSERT INTO minuta_neon_probe (payload) VALUES (:payload)"),
                {"payload": "rollback-test"},
            )
            trans.rollback()
        except Exception:
            trans.rollback()
            raise
        count = connection.scalar(text("SELECT COUNT(*) FROM minuta_neon_probe"))
        _LOGGER.info("rollback ok registros_probe=%s", count)


def _run_alembic() -> None:
    from auth.migration.alembic_runner import run_alembic_cli_upgrade

    start = time.perf_counter()
    run_alembic_cli_upgrade("head")
    elapsed_ms = round((time.perf_counter() - start) * 1000, 2)
    _LOGGER.info("alembic upgrade head ok duracao_ms=%s", elapsed_ms)


def _validate_schema(engine) -> None:
    inspector = inspect(engine)
    tables = set(inspector.get_table_names())
    missing = sorted(_EXPECTED_TABLES - tables)
    if missing:
        raise RuntimeError(f"Tabelas ausentes apos Alembic: {missing}")

    revision = None
    with engine.connect() as connection:
        if "alembic_version" in tables:
            revision = connection.scalar(text("SELECT version_num FROM alembic_version LIMIT 1"))
    _LOGGER.info("schema ok tabelas=%s alembic_revision=%s", len(tables), revision)

    for table_name in sorted(_EXPECTED_TABLES - {"alembic_version"}):
        fks = inspector.get_foreign_keys(table_name)
        indexes = inspector.get_indexes(table_name)
        _LOGGER.info(
            "schema.table table=%s fks=%s indexes=%s",
            table_name,
            len(fks),
            len(indexes),
        )


def _test_database_usage() -> None:
    from infrastructure.services.database_usage_service import DatabaseUsageService

    uso = DatabaseUsageService().medir()
    _LOGGER.info(
        "database_usage motor=%s bytes=%s limite=%s disponivel=%s percentual=%s",
        uso.motor,
        uso.bytes_ocupados,
        uso.bytes_limite,
        uso.bytes_disponiveis,
        uso.utilizacao_percentual,
    )
    assert uso.motor == "PostgreSQL"
    assert uso.bytes_ocupados is not None


def _benchmark_queries(engine) -> dict[str, float]:
    timings: dict[str, float] = {}
    queries = {
        "select_1": "SELECT 1",
        "count_usuario": "SELECT COUNT(*) FROM usuario",
        "count_configuracao": "SELECT COUNT(*) FROM configuracao",
        "pg_database_size": "SELECT pg_database_size(current_database())",
    }
    with engine.connect() as connection:
        for label, sql in queries.items():
            start = time.perf_counter()
            connection.scalar(text(sql))
            timings[label] = round((time.perf_counter() - start) * 1000, 2)
    return timings


def main() -> int:
    from core.settings import get_settings
    from infrastructure.database import configure_database, get_engine
    from infrastructure.schema import ensure_full_schema

    database_url = _require_postgres_url()
    _log_url_summary(database_url)

    get_settings.cache_clear()
    os.environ["MINUTA_DATABASE_URL"] = database_url
    get_settings.cache_clear()

    settings = get_settings()
    configure_database(
        database_url=settings.database_url,
        echo=settings.echo_sql,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    engine = get_engine()

    try:
        conn_info = _test_connection(engine)
        _LOGGER.info("conexao %s", conn_info)

        _test_read_write_rollback(engine)

        ensure_full_schema()
        _validate_schema(engine)
        _test_database_usage()

        query_timings = _benchmark_queries(engine)
        _LOGGER.info("benchmarks_ms %s", query_timings)

        _LOGGER.info("HOMOLOGACAO NEON N1/N2 CONCLUIDA COM SUCESSO")
        return 0
    finally:
        _reset_engine()
        get_settings.cache_clear()


if __name__ == "__main__":
    raise SystemExit(main())
