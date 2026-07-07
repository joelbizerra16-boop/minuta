from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Any

from sqlalchemy import MetaData, Table, event, select
from sqlalchemy.engine import Engine

from scripts.migration.constants import DOMAIN_TABLES, MIGRATION_TABLE_ORDER

_LOGGER = logging.getLogger("minuta.migration.extract")


@dataclass
class ExtractionResult:
    tables: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    row_counts: dict[str, int] = field(default_factory=dict)
    duration_ms: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "duration_ms": self.duration_ms,
            "row_counts": self.row_counts,
            "total_rows": sum(self.row_counts.values()),
        }


def create_readonly_sqlite_engine(database_url: str) -> Engine:
    from sqlalchemy import create_engine

    engine = create_engine(database_url, future=True)

    if engine.dialect.name == "sqlite":

        @event.listens_for(engine, "connect")
        def _set_query_only(dbapi_connection, _connection_record) -> None:  # type: ignore[no-untyped-def]
            dbapi_connection.execute("PRAGMA query_only=ON")

    return engine


def extract_all(engine: Engine) -> ExtractionResult:
    start = time.perf_counter()
    metadata = MetaData()
    result = ExtractionResult()

    for table_name in MIGRATION_TABLE_ORDER:
        if table_name not in DOMAIN_TABLES:
            continue
        if table_name not in inspect_table_names(engine):
            result.tables[table_name] = []
            result.row_counts[table_name] = 0
            continue

        table = Table(table_name, metadata, autoload_with=engine)
        with engine.connect() as connection:
            rows = [dict(row) for row in connection.execute(select(table)).mappings().all()]
        result.tables[table_name] = rows
        result.row_counts[table_name] = len(rows)
        _LOGGER.info("extract table=%s rows=%s", table_name, len(rows))

    result.duration_ms = round((time.perf_counter() - start) * 1000, 2)
    return result


def inspect_table_names(engine: Engine) -> set[str]:
    from sqlalchemy import inspect

    return set(inspect(engine).get_table_names())
