from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest
from sqlalchemy import text

from carregamentos.repository.sql_auditoria_nf_repository import _build_extrato_nf_sql
from carregamentos.repository.sql_rastreabilidade_nf_repository import _build_rastreabilidade_nf_sql
from infrastructure.database import configure_database, get_engine
from infrastructure.persistence.engine_info import get_engine_dialect


def _reset_engine() -> None:
    import infrastructure.database as db_module

    if db_module._engine is not None:
        db_module._engine.dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def test_extrato_nf_sql_postgresql_trim_syntax() -> None:
    sql = _build_extrato_nf_sql("postgresql")
    assert "TRIM(BOTH '0' FROM" in sql
    assert "TRIM(CAST(ic.numero_nf AS TEXT), '0')" not in sql


def test_rastreabilidade_sql_postgresql_json_agg() -> None:
    sql = _build_rastreabilidade_nf_sql("postgresql")
    assert "json_agg" in sql
    assert "json_group_array" not in sql


@pytest.mark.skipif(
    not os.getenv("MINUTA_TEST_POSTGRES_URL"),
    reason="Defina MINUTA_TEST_POSTGRES_URL para validar PostgreSQL/Neon.",
)
def test_postgresql_engine_connects_and_reports_dialect() -> None:
    postgres_url = os.environ["MINUTA_TEST_POSTGRES_URL"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_root = Path(tmp_dir)
        configure_database(
            database_url=postgres_url,
            data_root=data_root,
            pdf_storage_dir=data_root / "documentos",
            xml_storage_dir=data_root / "xml",
        )
        engine = get_engine()
        assert get_engine_dialect(engine) == "postgresql"
        with engine.connect() as connection:
            version = connection.scalar(text("SELECT version()"))
            assert version
        _reset_engine()
