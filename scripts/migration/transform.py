from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

from scripts.migration.constants import BOOLEAN_COLUMNS


def _to_bool(value: Any) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(int(value))
    text = str(value).strip().lower()
    if text in {"1", "true", "t", "yes"}:
        return True
    if text in {"0", "false", "f", "no"}:
        return False
    return bool(value)


def _normalize_value(column: str, value: Any, table_name: str) -> Any:
    if value is None:
        return None
    if column in BOOLEAN_COLUMNS.get(table_name, frozenset()):
        return _to_bool(value)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, datetime):
        return value
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, time):
        return value
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def transform_row(table_name: str, row: dict[str, Any]) -> dict[str, Any]:
    return {column: _normalize_value(column, value, table_name) for column, value in row.items()}


def transform_table(table_name: str, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    transformed = [transform_row(table_name, row) for row in rows]
    if table_name == "usuario":
        transformed.sort(key=lambda item: int(item.get("id") or 0))
    return transformed
