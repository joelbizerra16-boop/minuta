"""Expressoes SQL portaveis entre SQLite e PostgreSQL."""

from __future__ import annotations


def normalize_dialect(dialect_name: str) -> str:
    base = str(dialect_name or "").strip().lower()
    if "+" in base:
        base = base.split("+", 1)[0]
    return base


def trim_both_zeros(column_sql: str, *, dialect: str) -> str:
    """Remove zeros das extremidades de um identificador numerico de NF."""
    if normalize_dialect(dialect) == "postgresql":
        return f"TRIM(BOTH '0' FROM CAST({column_sql} AS TEXT))"
    return f"TRIM(CAST({column_sql} AS TEXT), '0')"


def boolean_is_true(column_sql: str) -> str:
    """Predicado booleano portavel (SQLite 3.39+ e PostgreSQL)."""
    return f"{column_sql} IS TRUE"


def json_array_subquery(
    *,
    dialect: str,
    alias: str,
    fields: list[tuple[str, str]],
    from_sql: str,
    order_by: str,
) -> str:
    """Subconsulta escalar que retorna array JSON de objetos."""
    object_pairs = ", ".join(f"'{key}', {expr}" for key, expr in fields)
    if normalize_dialect(dialect) == "postgresql":
        return f"""(
        SELECT COALESCE(
            json_agg(
                json_build_object({object_pairs})
                ORDER BY {order_by}
            ),
            '[]'::json
        )
        {from_sql}
    ) AS {alias}"""
    return f"""(
        SELECT json_group_array(
            json_object({object_pairs})
        )
        {from_sql}
        ORDER BY {order_by}
    ) AS {alias}"""
