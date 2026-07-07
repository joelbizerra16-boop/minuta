"""Utilitarios de persistencia (sem geracao manual de IDs)."""

from infrastructure.persistence.engine_info import (
    get_dialect_name,
    get_engine_dialect,
    get_sqlite_database_path,
    is_postgresql_engine,
    is_sqlite_engine,
)
from infrastructure.persistence.sql_compat import (
    boolean_is_true,
    json_array_subquery,
    normalize_dialect,
    trim_both_zeros,
)

__all__ = [
    "boolean_is_true",
    "get_dialect_name",
    "get_engine_dialect",
    "get_sqlite_database_path",
    "is_postgresql_engine",
    "is_sqlite_engine",
    "json_array_subquery",
    "normalize_dialect",
    "trim_both_zeros",
]
