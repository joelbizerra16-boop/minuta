"""Logs tecnicos do modulo de persistencia (sem credenciais)."""

from __future__ import annotations

import logging

from sqlalchemy.engine import Engine

from infrastructure.persistence.sql_compat import normalize_dialect

_LOGGER = logging.getLogger("minuta.persistence.bootstrap")


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
        strategy = "alembic_upgrade_head+create_all_checkfirst"
    else:
        strategy = "unsupported"
    _LOGGER.info(
        "persistence.schema_strategy dialect=%s strategy=%s",
        normalized,
        strategy,
    )
