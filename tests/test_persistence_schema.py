from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest

from core.bootstrap import configure_application_storage
from core.settings import get_settings
from infrastructure.database import get_engine
from infrastructure.persistence.engine_info import get_engine_dialect
from infrastructure.schema import ensure_full_schema


def test_ensure_full_schema_sqlite_uses_create_all() -> None:
    import infrastructure.database as db_module

    get_settings.cache_clear()
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "schema.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()
        configure_application_storage()
        assert get_engine_dialect(get_engine()) == "sqlite"
        assert db_path.exists()

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None
        get_settings.cache_clear()


@pytest.mark.skipif(
    not os.getenv("MINUTA_TEST_POSTGRES_URL"),
    reason="Defina MINUTA_TEST_POSTGRES_URL para validar PostgreSQL/Neon.",
)
def test_ensure_full_schema_postgresql_via_alembic() -> None:
    import infrastructure.database as db_module

    postgres_url = os.environ["MINUTA_TEST_POSTGRES_URL"]
    get_settings.cache_clear()
    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    get_settings.cache_clear()

    from infrastructure.database import configure_database

    settings = get_settings()
    configure_database(
        database_url=settings.database_url,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    ensure_full_schema()
    assert get_engine_dialect(get_engine()) == "postgresql"

    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None
    get_settings.cache_clear()
