from __future__ import annotations

import hashlib
import logging
import time
from dataclasses import dataclass, field
from typing import Any

from sqlalchemy import inspect, text
from sqlalchemy.engine import Engine

from scripts.migration.constants import DOMAIN_TABLES, FOREIGN_KEY_CHECKS

_LOGGER = logging.getLogger("minuta.migration.inventory")


@dataclass
class TableInventory:
    name: str
    row_count: int
    columns: list[str]
    primary_key: list[str]
    foreign_keys: list[dict[str, Any]]
    indexes: list[dict[str, Any]]
    checksum: str


@dataclass
class InventoryReport:
    tables: list[TableInventory] = field(default_factory=list)
    total_rows: int = 0
    duration_ms: float = 0.0
    source_dialect: str = "sqlite"

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_dialect": self.source_dialect,
            "total_rows": self.total_rows,
            "duration_ms": self.duration_ms,
            "tables": [
                {
                    "name": t.name,
                    "row_count": t.row_count,
                    "columns": t.columns,
                    "primary_key": t.primary_key,
                    "foreign_keys": t.foreign_keys,
                    "indexes": t.indexes,
                    "checksum": t.checksum,
                }
                for t in self.tables
            ],
        }


def _table_checksum(engine: Engine, table_name: str, pk_cols: list[str]) -> str:
    order_clause = ", ".join(pk_cols) if pk_cols else "rowid"
    query = text(f'SELECT * FROM "{table_name}" ORDER BY {order_clause}')
    digest = hashlib.sha256()
    with engine.connect() as connection:
        result = connection.execution_options(stream_results=True).execute(query)
        for row in result:
            digest.update(repr(tuple(row)).encode("utf-8"))
    return digest.hexdigest()


def build_inventory(engine: Engine) -> InventoryReport:
    start = time.perf_counter()
    inspector = inspect(engine)
    available = set(inspector.get_table_names())
    tables: list[TableInventory] = []
    total_rows = 0

    for table_name in DOMAIN_TABLES:
        if table_name not in available:
            tables.append(
                TableInventory(
                    name=table_name,
                    row_count=0,
                    columns=[],
                    primary_key=[],
                    foreign_keys=[],
                    indexes=[],
                    checksum="",
                )
            )
            continue

        pk = inspector.get_pk_constraint(table_name).get("constrained_columns") or []
        columns = [col["name"] for col in inspector.get_columns(table_name)]
        fks = inspector.get_foreign_keys(table_name)
        indexes = inspector.get_indexes(table_name)

        with engine.connect() as connection:
            count = int(connection.scalar(text(f'SELECT COUNT(*) FROM "{table_name}"')) or 0)

        checksum = _table_checksum(engine, table_name, pk) if count else ""
        total_rows += count
        tables.append(
            TableInventory(
                name=table_name,
                row_count=count,
                columns=columns,
                primary_key=list(pk),
                foreign_keys=fks,
                indexes=indexes,
                checksum=checksum,
            )
        )
        _LOGGER.info("inventory table=%s rows=%s", table_name, count)

    duration_ms = round((time.perf_counter() - start) * 1000, 2)
    return InventoryReport(
        tables=tables,
        total_rows=total_rows,
        duration_ms=duration_ms,
        source_dialect=engine.dialect.name,
    )


def validate_source_integrity(engine: Engine, inventory: InventoryReport) -> list[str]:
    """Validacao pre-carga: FK, nulos em PK, duplicidade de PK."""
    errors: list[str] = []
    row_maps: dict[str, dict[Any, set[Any]]] = {}

    with engine.connect() as connection:
        for table in inventory.tables:
            if table.row_count == 0:
                row_maps[table.name] = {}
                continue
            pk_cols = table.primary_key
            if not pk_cols:
                errors.append(f"pk_ausente:{table.name}")
                continue
            pk_col = pk_cols[0]
            rows = connection.execute(
                text(f'SELECT "{pk_col}" FROM "{table.name}"')
            ).fetchall()
            values = [row[0] for row in rows]
            if any(value is None for value in values):
                errors.append(f"pk_nulo:{table.name}")
            if len(values) != len(set(values)):
                errors.append(f"pk_duplicada:{table.name}")
            row_maps[table.name] = {pk_col: set(values)}

        for table_name, checks in FOREIGN_KEY_CHECKS.items():
            inv = next((t for t in inventory.tables if t.name == table_name), None)
            if inv is None or inv.row_count == 0:
                continue
            for child_col, parent_table, parent_col in checks:
                orphan_sql = text(
                    f'''
                    SELECT COUNT(*) FROM "{table_name}" child
                    WHERE child."{child_col}" IS NOT NULL
                      AND NOT EXISTS (
                        SELECT 1 FROM "{parent_table}" parent
                        WHERE parent."{parent_col}" = child."{child_col}"
                      )
                    '''
                )
                orphans = int(connection.scalar(orphan_sql) or 0)
                if orphans > 0:
                    errors.append(f"orfao:{table_name}.{child_col}->{parent_table}.{parent_col}:{orphans}")

    return errors
