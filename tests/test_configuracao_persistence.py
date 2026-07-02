from __future__ import annotations

import os
import tempfile
from pathlib import Path

from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.config_storage import CONFIG_CHAVE_SEPARACAO, SqlJsonConfigStorage


def test_configuracao_insert_sqlite() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'cfg.db').as_posix()}"
        get_settings.cache_clear()
        configure_database(
            database_url=os.environ["MINUTA_DATABASE_URL"],
            data_root=data_dir,
            pdf_storage_dir=data_dir / "documentos",
        )
        ensure_full_schema()
        storage = SqlJsonConfigStorage()
        storage.save_list(CONFIG_CHAVE_SEPARACAO, [{"NF": "1", "Produto": "Teste"}])
        loaded = storage.load_list(CONFIG_CHAVE_SEPARACAO, default=[])
        assert len(loaded) == 1
        assert loaded[0]["NF"] == "1"
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("configuracao sqlite insert OK")


if __name__ == "__main__":
    test_configuracao_insert_sqlite()
    print("All configuracao persistence tests passed")
