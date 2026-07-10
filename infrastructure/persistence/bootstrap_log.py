"""Logs tecnicos do modulo de persistencia (sem credenciais)."""

from __future__ import annotations

import logging

from sqlalchemy.engine import Engine

from core.database_config import DatabaseUrlSource, describe_database_target
from infrastructure.persistence.sql_compat import normalize_dialect

_LOGGER = logging.getLogger("minuta.persistence.bootstrap")


def log_database_resolution(
    *,
    runtime_environment: str,
    database_url_source: str,
    database_url: str,
) -> None:
    target = describe_database_target(database_url)
    _LOGGER.info(
        "persistence.database_resolution runtime=%s source=%s dialect=%s driver=%s host=%s database=%s schema=%s",
        runtime_environment,
        database_url_source,
        target["dialect"],
        target["driver"],
        target["host"] or "-",
        target["database"] or "-",
        target["schema"] or "-",
    )
    if database_url_source == DatabaseUrlSource.SQLITE_DEFAULT.value:
        _LOGGER.warning(
            "persistence.database_resolution sqlite_fallback_ativo "
            "runtime=%s database=%s",
            runtime_environment,
            target["database"],
        )


def log_engine_configured(engine: Engine) -> None:
    dialect = normalize_dialect(engine.dialect.name)
    driver = normalize_dialect(engine.url.drivername or "")
    masked_url = engine.url.render_as_string(hide_password=True)
    database = str(engine.url.database or "").strip() or ":memory:"
    _LOGGER.info(
        "persistence.engine_configured dialect=%s driver=%s database=%s url=%s",
        dialect,
        driver,
        database,
        masked_url,
    )


def log_schema_strategy(dialect: str) -> None:
    normalized = normalize_dialect(dialect)
    if normalized == "sqlite":
        strategy = "create_all"
    elif normalized == "postgresql":
        strategy = "alembic_upgrade_head"
    else:
        strategy = "unsupported"
    _LOGGER.info(
        "persistence.schema_strategy dialect=%s strategy=%s",
        normalized,
        strategy,
    )
