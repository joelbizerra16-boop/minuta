from __future__ import annotations

import logging

from infrastructure.models.base import Base as ModelBase
from infrastructure.persistence.bootstrap_log import log_schema_strategy
from infrastructure.persistence.engine_info import get_engine_dialect
from infrastructure.persistence.sqlite_schema_bootstrap import apply_sqlite_schema

_LOGGER = logging.getLogger("minuta.persistence.schema")


def ensure_full_schema() -> None:
    from infrastructure.database import get_engine

    engine = get_engine()
    dialect = get_engine_dialect(engine)
    log_schema_strategy(dialect)

    if dialect == "sqlite":
        _LOGGER.info("persistence.schema_apply method=create_all dialect=sqlite")
        apply_sqlite_schema(engine, ModelBase.metadata)
        return

    if dialect == "postgresql":
        from auth.migration.alembic_runner import run_alembic_cli_upgrade

        _LOGGER.info("persistence.schema_apply method=alembic_upgrade_head dialect=postgresql")
        run_alembic_cli_upgrade("head")
        return

    raise RuntimeError(f"Dialecto de banco nao suportado para bootstrap de schema: {dialect}")
