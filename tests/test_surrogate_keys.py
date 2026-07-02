"""Testes de geracao nativa de chaves primarias (sem MAX(id)+1)."""

from __future__ import annotations

import os
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from auth.models.usuario import Usuario
from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.security.password import hash_password
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.models.configuracao import ConfiguracaoORM
from infrastructure.repositories.configuracao_repository import ConfiguracaoRecord
from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
from infrastructure.schema import ensure_full_schema
from infrastructure.unit_of_work import UnitOfWork


def _configure_temp_sqlite(tmp_dir: Path) -> None:
    import infrastructure.database as db_module

    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_dir / 'pk.db').as_posix()}"
    get_settings.cache_clear()
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=tmp_dir,
        pdf_storage_dir=tmp_dir / "documentos",
    )
    ensure_full_schema()
    return db_module


def test_sqlite_autoincrement_configuracao() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        db_module = _configure_temp_sqlite(Path(tmp))
        repo = SqlConfiguracaoRepository()
        saved = repo.save(ConfiguracaoRecord(id=0, chave="teste.pk", valor='{"ok": true}'))
        assert saved.id > 0
        with UnitOfWork() as uow:
            row = uow.session.get(ConfiguracaoORM, saved.id)
            assert row is not None
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("sqlite autoincrement configuracao OK")


def test_concurrent_usuario_insert_distinct_ids() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        db_module = _configure_temp_sqlite(Path(tmp))
        repo = SqlUsuarioRepository()
        repo.ensure_default_admin("admin", "admin123", "Admin")
        errors: list[str] = []
        ids: list[int] = []
        lock = threading.Lock()

        def _create(index: int) -> None:
            try:
                saved = repo.save(
                    Usuario(
                        id=0,
                        nome=f"Operador {index}",
                        usuario=f"op{index}",
                        senha_hash=hash_password("senha123"),
                        perfil="OPERADOR",
                        ativo=True,
                        bloqueado=False,
                    )
                )
                with lock:
                    ids.append(saved.id)
            except Exception as exc:
                with lock:
                    errors.append(str(exc))

        with ThreadPoolExecutor(max_workers=8) as pool:
            futures = [pool.submit(_create, i) for i in range(1, 21)]
            for future in as_completed(futures):
                future.result()

        assert not errors, errors
        assert len(ids) == 20
        assert len(set(ids)) == 20
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        get_settings.cache_clear()
    print("concurrent usuario insert OK")


if __name__ == "__main__":
    test_sqlite_autoincrement_configuracao()
    test_concurrent_usuario_insert_distinct_ids()
    print("All surrogate key tests passed")
