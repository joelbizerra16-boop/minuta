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
_AUTH_STORAGE_CONFIGURED = False
_CONFIGURED_DATA_ROOT: Path | None = None


def reset_auth_bootstrap_state() -> None:
    """Utilitario de teste: permite reconfigurar auth storage."""
    global _repository, _AUTH_STORAGE_CONFIGURED, _CONFIGURED_DATA_ROOT
    _repository = None
    _AUTH_STORAGE_CONFIGURED = False
    _CONFIGURED_DATA_ROOT = None
    get_auth_service.cache_clear()
    get_usuario_service.cache_clear()


def configure_auth_storage(data_dir: Path) -> UsuarioRepository:
    global _repository, _AUTH_STORAGE_CONFIGURED, _CONFIGURED_DATA_ROOT
    if _AUTH_STORAGE_CONFIGURED and _CONFIGURED_DATA_ROOT == data_dir and _repository is not None:
        return _repository

    _ = data_dir
    repository = SqlUsuarioRepository()
    repository.ensure_default_admin(
        username=DEFAULT_ADMIN_USERNAME,
        password=DEFAULT_ADMIN_PASSWORD,
        nome=DEFAULT_ADMIN_NAME,
    )
    _repository = repository
    _AUTH_STORAGE_CONFIGURED = True
    _CONFIGURED_DATA_ROOT = data_dir
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
