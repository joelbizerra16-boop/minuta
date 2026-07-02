from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.repository.usuario_repository import UsuarioRepository
from auth.services.auth_service import AuthService
from auth.services.usuario_service import UsuarioService

DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"
DEFAULT_ADMIN_NAME = "Administrador"

_repository: UsuarioRepository | None = None


def configure_auth_storage(data_dir: Path) -> UsuarioRepository:
    global _repository
    _ = data_dir
    repository = SqlUsuarioRepository()
    repository.ensure_default_admin(
        username=DEFAULT_ADMIN_USERNAME,
        password=DEFAULT_ADMIN_PASSWORD,
        nome=DEFAULT_ADMIN_NAME,
    )
    _repository = repository
    return _repository


def get_usuario_repository() -> UsuarioRepository:
    if _repository is None:
        raise RuntimeError("Auth storage not configured. Call configure_auth_storage first.")
    return _repository


@lru_cache(maxsize=1)
def get_auth_service() -> AuthService:
    return AuthService(get_usuario_repository())


@lru_cache(maxsize=1)
def get_usuario_service() -> UsuarioService:
    return UsuarioService(get_usuario_repository())
