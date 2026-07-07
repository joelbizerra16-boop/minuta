from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from typing import Any

from sqlalchemy import MetaData, Table, insert, text
from sqlalchemy.engine import Connection, Engine

from scripts.migration.constants import (
    DEFAULT_BATCH_SIZE,
    MIGRATION_TABLE_ORDER,
    TRUNCATE_TABLE_ORDER,
)
from scripts.migration.transform import transform_table

_LOGGER = logging.getLogger("minuta.migration.load")


@dataclass
class LoadReport:
    tables_loaded: dict[str, int] = field(default_factory=dict)
    duration_ms: float = 0.0
    rolled_back: bool = False
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "tables_loaded": self.tables_loaded,
            "total_rows": sum(self.tables_loaded.values()),
            "duration_ms": self.duration_ms,
            "rolled_back": self.rolled_back,
            "error": self.error,
        }


def _truncate_target(connection: Connection) -> None:
    table_list = ", ".join(f'"{name}"' for name in TRUNCATE_TABLE_ORDER)
    connection.execute(text(f"TRUNCATE TABLE {table_list} RESTART IDENTITY CASCADE"))


def _bulk_insert(
    connection: Connection,
    table: Table,
    rows: list[dict[str, Any]],
    batch_size: int,
) -> int:
    if not rows:
        return 0
    inserted = 0
    for offset in range(0, len(rows), batch_size):
        batch = rows[offset : offset + batch_size]
        connection.execute(insert(table), batch)
        inserted += len(batch)
    return inserted


def _sync_sequences(connection: Connection, metadata: MetaData) -> None:
    for table_name in MIGRATION_TABLE_ORDER:
        table = metadata.tables.get(table_name)
        if table is None or "id" not in table.c:
            continue
        connection.execute(
            text(
                f'''
                SELECT setval(
                    pg_get_serial_sequence('{table_name}', 'id'),
                    COALESCE((SELECT MAX(id) FROM "{table_name}"), 1),
                    (SELECT COUNT(*) > 0 FROM "{table_name}")
                )
                '''
            )
        )


def load_dataset(
    target_engine: Engine,
    dataset: dict[str, list[dict[str, Any]]],
    *,
    batch_size: int = DEFAULT_BATCH_SIZE,
    truncate: bool = True,
) -> LoadReport:
    start = time.perf_counter()
    metadata = MetaData()
    report = LoadReport()

    try:
        with target_engine.begin() as connection:
            if truncate:
                _truncate_target(connection)

            for table_name in MIGRATION_TABLE_ORDER:
                rows = transform_table(table_name, dataset.get(table_name, []))
                if table_name not in metadata.tables:
                    Table(table_name, metadata, autoload_with=connection)
                table = metadata.tables[table_name]
                loaded = _bulk_insert(connection, table, rows, batch_size)
                report.tables_loaded[table_name] = loaded
                _LOGGER.info("load table=%s rows=%s", table_name, loaded)

            _sync_sequences(connection, metadata)
    except Exception as exc:
        report.rolled_back = True
        report.error = f"{type(exc).__name__}: {exc}"
        _LOGGER.exception("load falhou — rollback automatico")
        raise
    finally:
        report.duration_ms = round((time.perf_counter() - start) * 1000, 2)

    return report
