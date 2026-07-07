from __future__ import annotations

import os
import tempfile
from pathlib import Path

from infrastructure.database import configure_database, get_engine
from infrastructure.services.database_usage_service import DatabaseUsageService
from core.retention_policy import DATABASE_STORAGE_LIMIT_BYTES


def _configure_sqlite(db_path: Path) -> None:
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=db_path.parent,
        pdf_storage_dir=db_path.parent / "documentos",
        xml_storage_dir=db_path.parent / "xml",
    )
    get_engine().dispose()


def test_medir_sqlite_por_arquivo_do_engine() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "uso.db"
        _configure_sqlite(db_path)
        db_path.write_bytes(b"x" * 2048)

        uso = DatabaseUsageService().medir()

        assert uso.motor == "SQLite"
        assert uso.bytes_ocupados == 2048
        assert uso.bytes_limite == DATABASE_STORAGE_LIMIT_BYTES
        assert uso.bytes_disponiveis == DATABASE_STORAGE_LIMIT_BYTES - 2048
        assert uso.utilizacao_percentual is not None
        assert uso.utilizacao_percentual >= 0
        assert uso.bytes_ocupados is not None

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


def test_medir_sqlite_fallback_pragma_em_memoria() -> None:
    import infrastructure.database as db_module

    configure_database(
        database_url="sqlite:///:memory:",
        data_root=Path(tempfile.gettempdir()),
        pdf_storage_dir=Path(tempfile.gettempdir()) / "documentos",
        xml_storage_dir=Path(tempfile.gettempdir()) / "xml",
    )

    with get_engine().connect() as connection:
        connection.exec_driver_sql(
            "CREATE TABLE IF NOT EXISTS medicao_teste (id INTEGER PRIMARY KEY, payload TEXT)"
        )
        connection.exec_driver_sql(
            "INSERT INTO medicao_teste (id, payload) VALUES (1, '"
            + ("x" * 4096)
            + "')"
        )
        connection.commit()

    uso = DatabaseUsageService().medir()

    assert uso.motor == "SQLite"
    assert uso.bytes_ocupados is not None
    assert uso.bytes_ocupados > 0
    assert "PRAGMA" in str(uso.observacao or "")

    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def test_medir_sqlite_caminho_relativo_via_engine() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        original_cwd = Path.cwd()
        os.chdir(tmp_dir)
        try:
            db_path = Path("data") / "relativo.db"
            db_path.parent.mkdir(parents=True, exist_ok=True)
            _configure_sqlite(db_path)
            db_path.write_bytes(b"y" * 1024)

            uso = DatabaseUsageService().medir()

            assert uso.motor == "SQLite"
            assert uso.bytes_ocupados == 1024
            assert uso.bytes_ocupados is not None
        finally:
            os.chdir(original_cwd)
            get_engine().dispose()
            db_module._engine = None
            db_module._session_factory = None
            db_module._data_root = None
            db_module._pdf_storage_dir = None
            db_module._xml_storage_dir = None


if __name__ == "__main__":
    test_medir_sqlite_por_arquivo_do_engine()
    test_medir_sqlite_fallback_pragma_em_memoria()
    test_medir_sqlite_caminho_relativo_via_engine()
    print("test_database_usage_service OK")
