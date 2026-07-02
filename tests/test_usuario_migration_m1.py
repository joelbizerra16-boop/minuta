#!/usr/bin/env python
"""Testes da migracao M1 de usuarios."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from auth.bootstrap import configure_auth_storage
from auth.migration.alembic_runner import run_alembic_downgrade, run_alembic_upgrade
from auth.migration.migrate_usuarios import backup_usuarios_json, migrate_usuarios_from_json
from auth.migration.usuario_comparator import compare_usuarios
from auth.repository.dual_usuario_repository import DualUsuarioRepository
from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.repository.usuario_repository import JsonUsuarioRepository
from auth.services.auth_service import AuthService
from auth.services.usuario_service import UsuarioService
from core.settings import StorageBackend, get_settings
from infrastructure.database import configure_database, get_engine


def _configure_temp_db(tmp_dir: Path) -> str:
    db_path = tmp_dir / "m1_test.db"
    url = f"sqlite:///{db_path.as_posix()}"
    configure_database(
        database_url=url,
        data_root=tmp_dir,
        pdf_storage_dir=tmp_dir / "documentos",
    )
    return url


def test_sql_usuario_repository_crud() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        url = _configure_temp_db(tmp_dir)
        run_alembic_upgrade("head")

        json_repo = JsonUsuarioRepository(tmp_dir / "usuarios.json")
        json_repo.ensure_default_admin("admin", "admin123", "Administrador")
        sql_repo = SqlUsuarioRepository()

        json_users = json_repo.list_all(include_inactive=True)
        for user in json_users:
            sql_repo.save(user)

        sql_users = sql_repo.list_all(include_inactive=True)
        report = compare_usuarios(json_users, sql_users)
        assert report.success

        run_alembic_downgrade("base")
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("sql usuario repository CRUD OK")


def test_migration_script_with_backup_and_rollback() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        json_path = tmp_dir / "usuarios.json"
        json_repo = JsonUsuarioRepository(json_path)
        json_repo.ensure_default_admin("admin", "admin123", "Administrador")
        original = json_path.read_bytes()

        backup = backup_usuarios_json(json_path, tmp_dir / "backups")
        assert backup.is_file()

        url = _configure_temp_db(tmp_dir)
        report = migrate_usuarios_from_json(
            json_path,
            database_url=url,
            data_root=tmp_dir,
            pdf_storage_dir=tmp_dir / "documentos",
            create_backup=False,
        )
        assert report.success, report.to_text()
        assert report.migrated_count == 1
        assert json_path.read_bytes() == original

        run_alembic_downgrade("base")
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("migration script backup and validation OK")


def test_dual_repository_json_official() -> None:
    import infrastructure.database as db_module

    from auth.repository.dual_usuario_repository import DualUsuarioRepository
    from auth.repository.usuario_repository import JsonUsuarioRepository
    from auth.repository.sql_usuario_repository import SqlUsuarioRepository

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        os.environ["MINUTA_DATA_ROOT"] = str(tmp_dir)
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_dir / 'dual.db').as_posix()}"
        get_settings.cache_clear()
        configure_database(
            database_url=os.environ["MINUTA_DATABASE_URL"],
            data_root=tmp_dir,
            pdf_storage_dir=tmp_dir / "documentos",
        )
        from infrastructure.schema import ensure_full_schema

        ensure_full_schema()

        sql_repo = SqlUsuarioRepository()
        repo = DualUsuarioRepository(tmp_dir / "usuarios.json", sql_repo)
        repo.ensure_default_admin("admin", "admin123", "Administrador")

        auth = AuthService(repo)
        users = UsuarioService(repo)

        assert auth.authenticate("admin", "admin123").success
        operator = users.create_user("Operador Dual", "dualop", "dual1234", "OPERADOR")
        assert operator.usuario == "dualop"
        users.block_user(operator.id)
        assert auth.authenticate("dualop", "dual1234").blocked
        users.unblock_user(operator.id)
        users.delete_user(operator.id)
        assert not auth.authenticate("dualop", "dual1234").success

        payload = json.loads((tmp_dir / "usuarios.json").read_text(encoding="utf-8"))
        assert len(payload["usuarios"]) == 2

        get_settings.cache_clear()
        os.environ.pop("MINUTA_STORAGE_BACKEND", None)
        os.environ.pop("MINUTA_DATA_ROOT", None)
        os.environ.pop("MINUTA_DATABASE_URL", None)

        run_alembic_downgrade("base")
        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("dual repository json official OK")


def test_json_mode_unchanged() -> None:
    os.environ.pop("MINUTA_STORAGE_BACKEND", None)
    get_settings.cache_clear()
    settings = get_settings()
    assert settings.storage_backend == StorageBackend.SQL
    print("sql mode default OK")


if __name__ == "__main__":
    test_json_mode_unchanged()
    test_sql_usuario_repository_crud()
    test_migration_script_with_backup_and_rollback()
    test_dual_repository_json_official()
    print("All M1 usuario migration tests passed.")
