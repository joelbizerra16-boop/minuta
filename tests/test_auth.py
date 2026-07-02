#!/usr/bin/env python
"""Testes locais do modulo de autenticacao."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from auth.bootstrap import configure_auth_storage
from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.security.password import hash_password, verify_password
from auth.services.auth_service import AuthService
from auth.services.usuario_service import UsuarioService
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine


def _setup_temp_sql_env(tmp_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(tmp_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_dir / 'auth.db').as_posix()}"
    get_settings.cache_clear()
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=tmp_dir,
        pdf_storage_dir=tmp_dir / "documentos",
    )
    from infrastructure.schema import ensure_full_schema

    ensure_full_schema()


def test_password_hashing() -> None:
    stored = hash_password("admin123")
    assert stored.startswith("pbkdf2_sha256$")
    assert verify_password("admin123", stored)
    assert not verify_password("wrong", stored)
    print("password hashing OK")


def test_auth_flow() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_temp_sql_env(data_dir)
        configure_auth_storage(data_dir)
        repo = SqlUsuarioRepository()
        auth = AuthService(repo)
        users = UsuarioService(repo)

        valid = auth.authenticate("admin", "admin123")
        assert valid.success and valid.user is not None

        invalid = auth.authenticate("admin", "wrong")
        assert not invalid.success

        operator = users.create_user("Operador Teste", "operador1", "oper123", "OPERADOR")
        assert operator.usuario == "operador1"

        op_login = auth.authenticate("operador1", "oper123")
        assert op_login.success

        users.block_user(operator.id)
        blocked = auth.authenticate("operador1", "oper123")
        assert not blocked.success and blocked.blocked

        users.unblock_user(operator.id)
        assert auth.authenticate("operador1", "oper123").success

        users.change_password(operator.id, "nova1234")
        assert auth.authenticate("operador1", "nova1234").success

        users.delete_user(operator.id)
        deleted = auth.authenticate("operador1", "nova1234")
        assert not deleted.success

        persisted = repo.get_by_username("admin")
        assert persisted is not None
        assert persisted.senha_hash.startswith("pbkdf2_sha256$")
        assert "admin123" not in persisted.senha_hash
        print("auth flow OK")

        import infrastructure.database as db_module

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None


if __name__ == "__main__":
    test_password_hashing()
    test_auth_flow()
    print("All auth tests passed")
