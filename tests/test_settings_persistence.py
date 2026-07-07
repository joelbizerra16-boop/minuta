from __future__ import annotations

import os
import tempfile
from pathlib import Path

from core.settings import derive_storage_paths, get_settings


def test_derive_storage_paths_from_sqlite_url() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "data" / "minuta.db"
        data_root, pdf_dir, xml_dir = derive_storage_paths(f"sqlite:///{db_path.as_posix()}")
        assert data_root == db_path.parent.resolve()
        assert pdf_dir == data_root / "documentos"
        assert xml_dir == data_root / "xml_storage"


def test_derive_storage_paths_from_postgresql_url() -> None:
    data_root, pdf_dir, xml_dir = derive_storage_paths("postgresql://user:pass@localhost:5432/minuta")
    assert data_root == (Path.cwd() / "data").resolve()
    assert pdf_dir == data_root / "documentos"
    assert xml_dir == data_root / "xml_storage"


def test_default_database_url_is_relative_to_cwd() -> None:
    get_settings.cache_clear()
    os.environ.pop("MINUTA_DATABASE_URL", None)
    settings = get_settings()
    assert settings.database_url.startswith("sqlite:///")
    assert "minuta.db" in settings.database_url
    get_settings.cache_clear()
