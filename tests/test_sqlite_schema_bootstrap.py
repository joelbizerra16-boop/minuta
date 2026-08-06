from __future__ import annotations

import sqlite3
import tempfile
import threading
from contextlib import contextmanager
from collections.abc import Iterator
from pathlib import Path

import pytest
from sqlalchemy import inspect

from infrastructure.database import configure_database, get_engine
from infrastructure.models import Base
from infrastructure.persistence.sqlite_schema_bootstrap import (
    apply_sqlite_schema,
    reset_sqlite_schema_bootstrap_state,
)
from infrastructure.schema import ensure_full_schema


def _dispose_configured_engine() -> None:
    import infrastructure.database as db_module

    reset_sqlite_schema_bootstrap_state()
    engine = db_module._engine
    if engine is not None:
        engine.dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


@pytest.fixture(autouse=True)
def _reset_sqlite_bootstrap_state() -> None:
    yield
    _dispose_configured_engine()


@contextmanager
def _sqlite_test_database(name: str) -> Iterator[Path]:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / name
        _configure_sqlite(db_path)
        try:
            yield db_path
        finally:
            _dispose_configured_engine()


def _sqlite_url(db_path: Path) -> str:
    return f"sqlite:///{db_path.as_posix()}"


def _configure_sqlite(db_path: Path) -> None:
    reset_sqlite_schema_bootstrap_state()
    configure_database(
        database_url=_sqlite_url(db_path),
        data_root=db_path.parent,
        pdf_storage_dir=db_path.parent / "documentos",
        xml_storage_dir=db_path.parent / "xml_storage",
    )


def _list_tables(db_path: Path) -> list[str]:
    connection = sqlite3.connect(db_path)
    try:
        return [
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY 1"
            )
        ]
    finally:
        connection.close()


def test_sqlite_bootstrap_empty_database() -> None:
    with _sqlite_test_database("empty.db") as db_path:
        ensure_full_schema()

        tables = set(_list_tables(db_path))
        assert tables == set(Base.metadata.tables.keys())
        assert db_path.exists()


def test_sqlite_bootstrap_existing_database() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "existing.db"
        connection = sqlite3.connect(db_path)
        connection.execute(
            """
            CREATE TABLE motorista (
                id INTEGER PRIMARY KEY,
                nome VARCHAR(200) NOT NULL,
                criado_em DATETIME,
                atualizado_em DATETIME,
                ativo BOOLEAN DEFAULT 1,
                excluido_em DATETIME
            )
            """
        )
        connection.commit()
        connection.close()

        _configure_sqlite(db_path)
        try:
            ensure_full_schema()

            tables = set(_list_tables(db_path))
            assert tables == set(Base.metadata.tables.keys())
        finally:
            _dispose_configured_engine()


def test_sqlite_bootstrap_adds_xml_content_to_legacy_table() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "legacy_documento_xml.db"
        connection = sqlite3.connect(db_path)
        connection.execute(
            """
            CREATE TABLE documento_xml (
                id INTEGER PRIMARY KEY,
                chave_nfe CHAR(44) NOT NULL,
                numero_nf VARCHAR(20) NOT NULL,
                nome_arquivo VARCHAR(255) NOT NULL,
                caminho_arquivo VARCHAR(500) NOT NULL,
                hash_sha256 VARCHAR(64) NOT NULL,
                tamanho INTEGER NOT NULL DEFAULT 0,
                usuario_id INTEGER,
                data_importacao DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                ativo BOOLEAN NOT NULL DEFAULT 1
            )
            """
        )
        connection.commit()
        connection.close()

        _configure_sqlite(db_path)
        try:
            ensure_full_schema()
            columns = {column["name"] for column in inspect(get_engine()).get_columns("documento_xml")}
            assert "conteudo_xml" in columns
        finally:
            _dispose_configured_engine()


def test_sqlite_bootstrap_consecutive_calls_are_idempotent() -> None:
    with _sqlite_test_database("consecutive.db") as db_path:
        for _ in range(3):
            ensure_full_schema()

        tables = set(_list_tables(db_path))
        assert tables == set(Base.metadata.tables.keys())
        inspector = inspect(get_engine())
        assert set(inspector.get_table_names()) == tables


def test_sqlite_bootstrap_concurrent_calls_do_not_raise() -> None:
    with _sqlite_test_database("concurrent.db") as db_path:
        engine = get_engine()
        errors: list[tuple[str, str]] = []
        barrier = threading.Barrier(4)

        def worker(name: str) -> None:
            barrier.wait()
            try:
                apply_sqlite_schema(engine, Base.metadata)
            except Exception as exc:
                errors.append((name, f"{type(exc).__name__}: {exc}"))

        threads = [threading.Thread(target=worker, args=(f"t{index}",)) for index in range(4)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()

        assert errors == [], errors
        tables = set(_list_tables(db_path))
        assert tables == set(Base.metadata.tables.keys())


def test_sqlite_bootstrap_does_not_mask_failures() -> None:
    with _sqlite_test_database("invalid.db") as db_path:
        engine = get_engine()

        original_create_all = Base.metadata.create_all

        def _boom(*args, **kwargs):  # type: ignore[no-untyped-def]
            raise RuntimeError("bootstrap_failure_injected")

        Base.metadata.create_all = _boom  # type: ignore[method-assign]
        try:
            with pytest.raises(RuntimeError, match="bootstrap_failure_injected"):
                apply_sqlite_schema(engine, Base.metadata)
            reset_sqlite_schema_bootstrap_state()
            with pytest.raises(RuntimeError, match="bootstrap_failure_injected"):
                apply_sqlite_schema(engine, Base.metadata)
        finally:
            Base.metadata.create_all = original_create_all  # type: ignore[method-assign]

        assert _list_tables(db_path) == []
