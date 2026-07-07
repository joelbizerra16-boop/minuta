from __future__ import annotations

from infrastructure.persistence.sql_compat import (
    boolean_is_true,
    json_array_subquery,
    normalize_dialect,
    trim_both_zeros,
)


def test_normalize_dialect_strips_driver_suffix() -> None:
    assert normalize_dialect("postgresql+psycopg2") == "postgresql"
    assert normalize_dialect("sqlite") == "sqlite"


def test_trim_both_zeros_sqlite_syntax() -> None:
    expr = trim_both_zeros("ic.numero_nf", dialect="sqlite")
    assert expr == "TRIM(CAST(ic.numero_nf AS TEXT), '0')"


def test_trim_both_zeros_postgresql_syntax() -> None:
    expr = trim_both_zeros("ic.numero_nf", dialect="postgresql")
    assert expr == "TRIM(BOTH '0' FROM CAST(ic.numero_nf AS TEXT))"


def test_boolean_is_true_portable() -> None:
    assert boolean_is_true("dx.ativo") == "dx.ativo IS TRUE"


def test_json_array_subquery_sqlite_uses_json_group_array() -> None:
    sql = json_array_subquery(
        dialect="sqlite",
        alias="payload",
        fields=[("evento", "ho.evento")],
        from_sql="FROM historico_operacional ho",
        order_by="ho.id",
    )
    assert "json_group_array" in sql
    assert "json_object" in sql


def test_json_array_subquery_postgresql_uses_json_agg() -> None:
    sql = json_array_subquery(
        dialect="postgresql",
        alias="payload",
        fields=[("evento", "ho.evento")],
        from_sql="FROM historico_operacional ho",
        order_by="ho.id",
    )
    assert "json_agg" in sql
    assert "json_build_object" in sql
    assert "'[]'::json" in sql
