from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest

from core.bootstrap import configure_application_storage
from auth.bootstrap import configure_auth_storage, get_auth_service
from infrastructure.database import configure_database, get_engine
from infrastructure.models import Base
from infrastructure.schema import ensure_full_schema


def _seed_sqlite_database(db_path: Path) -> None:
    data_root = db_path.parent
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=data_root,
        pdf_storage_dir=data_root / "documentos",
        xml_storage_dir=data_root / "xml_storage",
    )
    ensure_full_schema()
    configure_auth_storage(data_root)
    engine = get_engine()
    Base.metadata.create_all(engine, checkfirst=True)
    configure_auth_storage(data_root)
    assert get_auth_service().authenticate("admin", "admin123").success

    import infrastructure.database as db_module

    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def test_migration_inventory_and_dry_run_on_seeded_sqlite() -> None:
    from scripts.migration.runner import run_migration

    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "source.db"
        _seed_sqlite_database(db_path)
        os.environ["MINUTA_SQLITE_SOURCE_URL"] = f"sqlite:///{db_path.as_posix()}"
        os.environ.pop("MINUTA_DATABASE_URL", None)

        report = run_migration(dry_run=True)
        assert report.aprovada
        assert report.validacao_pre_carga["ok"]
        assert report.inventario["total_rows"] >= 2  # perfil seed + admin usuario
        assert report.extracao["total_rows"] >= 2

        os.environ.pop("MINUTA_SQLITE_SOURCE_URL", None)


@pytest.mark.skipif(
    not os.getenv("MINUTA_TEST_POSTGRES_URL"),
    reason="Defina MINUTA_TEST_POSTGRES_URL para validar migracao completa.",
)
def test_migration_sqlite_to_postgresql_roundtrip() -> None:
    from scripts.migration.runner import run_migration

    postgres_url = os.environ["MINUTA_TEST_POSTGRES_URL"]
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "source.db"
        _seed_sqlite_database(db_path)
        os.environ["MINUTA_SQLITE_SOURCE_URL"] = f"sqlite:///{db_path.as_posix()}"
        os.environ["MINUTA_DATABASE_URL"] = postgres_url

        report = run_migration(dry_run=False)
        assert report.aprovada, report.bloqueador
        assert report.validacao_pos_carga["equivalente"]
        assert report.carga["total_rows"] >= 2

        os.environ.pop("MINUTA_SQLITE_SOURCE_URL", None)
        os.environ.pop("MINUTA_DATABASE_URL", None)
