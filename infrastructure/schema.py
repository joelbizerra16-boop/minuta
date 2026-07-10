from __future__ import annotations

import logging
import threading

from infrastructure.models.base import Base as ModelBase
from infrastructure.persistence.bootstrap_log import log_schema_strategy
from infrastructure.persistence.engine_info import get_engine_dialect
from infrastructure.persistence.sqlite_schema_bootstrap import apply_sqlite_schema

_LOGGER = logging.getLogger("minuta.persistence.schema")

_SCHEMA_LOCK = threading.Lock()
_SCHEMA_VERIFIED_KEY: str | None = None
_ALEMBIC_RUN_COUNT = 0


def get_alembic_run_count() -> int:
    return _ALEMBIC_RUN_COUNT


def reset_schema_bootstrap_state() -> None:
    """Utilitario de teste: permite reexecutar bootstrap de schema."""
    global _SCHEMA_VERIFIED_KEY, _ALEMBIC_RUN_COUNT
    with _SCHEMA_LOCK:
        _SCHEMA_VERIFIED_KEY = None
        _ALEMBIC_RUN_COUNT = 0


def _schema_key_for_engine(engine) -> str:
    return engine.url.render_as_string(hide_password=False)


def ensure_full_schema() -> None:
    from infrastructure.database import get_engine

    global _SCHEMA_VERIFIED_KEY, _ALEMBIC_RUN_COUNT

    engine = get_engine()
    dialect = get_engine_dialect(engine)
    schema_key = _schema_key_for_engine(engine)

    if _SCHEMA_VERIFIED_KEY == schema_key:
        _LOGGER.debug("persistence.schema skipped schema_key=%s dialect=%s", schema_key, dialect)
        return

    with _SCHEMA_LOCK:
        if _SCHEMA_VERIFIED_KEY == schema_key:
            _LOGGER.debug(
                "persistence.schema skipped_after_lock schema_key=%s dialect=%s",
                schema_key,
                dialect,
            )
            return

        log_schema_strategy(dialect)

        if dialect == "sqlite":
            _LOGGER.info("persistence.schema_apply method=create_all dialect=sqlite")
            apply_sqlite_schema(engine, ModelBase.metadata)
        elif dialect == "postgresql":
            from auth.migration.alembic_runner import run_alembic_cli_upgrade

            _LOGGER.info("persistence.schema_apply method=alembic_upgrade_head dialect=postgresql")
            run_alembic_cli_upgrade("head")
            _ALEMBIC_RUN_COUNT += 1
        else:
            raise RuntimeError(f"Dialecto de banco nao suportado para bootstrap de schema: {dialect}")

        _SCHEMA_VERIFIED_KEY = schema_key
        _LOGGER.info("persistence.schema verified schema_key=%s dialect=%s", schema_key, dialect)
