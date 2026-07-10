"""Carregamento do arquivo .env e secrets do runtime (Streamlit Cloud)."""

from __future__ import annotations

import logging
import os
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

_LOGGER = logging.getLogger("minuta.config.env")

_STREAMLIT_PROMOTED_KEYS: set[str] = set()
_DATABASE_URL_ENV_KEY = "MINUTA_DATABASE_URL"

# Caminhos aninhados suportados em secrets.toml (Streamlit Cloud / Neon).
_NESTED_DATABASE_URL_PATHS: tuple[tuple[str, ...], ...] = (
    ("minuta", "database_url"),
    ("neon", "database_url"),
    ("database", "url"),
    ("connections", "postgresql", "url"),
    ("connections", "neon", "url"),
    ("connections", "neon", "database_url"),
)


@dataclass(frozen=True)
class DotenvLoadResult:
    found: bool
    loaded: bool
    path: Path
    message: str


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def reset_streamlit_secret_state() -> None:
    _STREAMLIT_PROMOTED_KEYS.clear()


def streamlit_promoted_env_keys() -> frozenset[str]:
    return frozenset(_STREAMLIT_PROMOTED_KEYS)


def _is_scalar_secret(value: object) -> bool:
    return isinstance(value, (str, int, float))


def _lookup_nested_secret(store: Mapping[str, object], path: tuple[str, ...]) -> str | None:
    current: object = store
    for segment in path:
        if not isinstance(current, Mapping):
            return None
        if segment not in current:
            return None
        current = current[segment]
    if not _is_scalar_secret(current):
        return None
    stripped = str(current).strip()
    return stripped or None


def _nested_path_label(path: tuple[str, ...]) -> str:
    return ".".join(path)


def _promote_env_var(name: str, value: str, *, source_path: str) -> None:
    if os.getenv(name):
        return
    os.environ[name] = value
    _STREAMLIT_PROMOTED_KEYS.add(name)
    _LOGGER.info(
        "config.secrets origin=streamlit action=promote target_key=%s source_path=%s status=found",
        name,
        source_path,
    )


def _promote_streamlit_scalar_secrets(store: Mapping[str, object]) -> None:
    for key, value in store.items():
        if _is_scalar_secret(value):
            _promote_env_var(str(key), str(value), source_path=str(key))


def _resolve_database_url_from_secrets(store: Mapping[str, object]) -> tuple[str, str] | None:
    root = store.get(_DATABASE_URL_ENV_KEY)
    if _is_scalar_secret(root):
        stripped = str(root).strip()
        if stripped:
            return stripped, _DATABASE_URL_ENV_KEY

    for path in _NESTED_DATABASE_URL_PATHS:
        resolved = _lookup_nested_secret(store, path)
        if resolved:
            return resolved, _nested_path_label(path)
    return None


def _promote_database_url_from_nested_secrets(store: Mapping[str, object]) -> None:
    if os.getenv(_DATABASE_URL_ENV_KEY):
        return
    resolved = _resolve_database_url_from_secrets(store)
    if resolved is None:
        _LOGGER.info(
            "config.secrets origin=streamlit action=promote target_key=%s status=absent",
            _DATABASE_URL_ENV_KEY,
        )
        return
    value, source_path = resolved
    _promote_env_var(_DATABASE_URL_ENV_KEY, value, source_path=source_path)


def _read_streamlit_secrets_mapping() -> Mapping[str, Any] | None:
    try:
        import streamlit as st
    except ImportError:
        _LOGGER.debug(
            "config.secrets origin=streamlit status=skipped reason=streamlit_import_error",
        )
        return None

    try:
        from streamlit.errors import StreamlitSecretNotFoundError
    except ImportError:
        StreamlitSecretNotFoundError = type(  # type: ignore[misc, assignment]
            "StreamlitSecretNotFoundError",
            (Exception,),
            {},
        )

    try:
        secrets_store = st.secrets
        if hasattr(secrets_store, "to_dict"):
            secrets_mapping: Mapping[str, Any] = secrets_store.to_dict()
        else:
            secrets_mapping = dict(secrets_store)
    except StreamlitSecretNotFoundError:
        _LOGGER.info(
            "config.secrets origin=streamlit status=not_found reason=StreamlitSecretNotFoundError",
        )
        return None
    except (KeyError, TypeError, AttributeError) as exc:
        _LOGGER.warning(
            "config.secrets origin=streamlit status=access_error reason=%s",
            type(exc).__name__,
        )
        return None

    if not secrets_mapping:
        _LOGGER.info("config.secrets origin=streamlit status=empty")
        return None

    return secrets_mapping


def hydrate_runtime_secrets() -> bool:
    """
    Hidrata os.environ a partir de st.secrets antes da resolucao de configuracao.

    No Streamlit Cloud, MINUTA_DATABASE_URL costuma existir em secrets.toml.
    Esta funcao garante a hidratacao antes do bootstrap e cobre layouts aninhados
    oficiais do projeto, incluindo [connections.neon] do Streamlit Cloud.
    """
    secrets_mapping = _read_streamlit_secrets_mapping()
    if secrets_mapping is None:
        return False

    _promote_streamlit_scalar_secrets(secrets_mapping)
    _promote_database_url_from_nested_secrets(secrets_mapping)
    return bool(_STREAMLIT_PROMOTED_KEYS)


def lookup_runtime_secret(name: str) -> str | None:
    """Leitura direta de st.secrets para chaves escalares (fallback de get_env)."""
    secrets_mapping = _read_streamlit_secrets_mapping()
    if secrets_mapping is None:
        return None

    if name in secrets_mapping and _is_scalar_secret(secrets_mapping[name]):
        stripped = str(secrets_mapping[name]).strip()
        if stripped:
            _LOGGER.info(
                "config.secrets origin=streamlit action=lookup key=%s source_path=%s status=found",
                name,
                name,
            )
            return stripped

    if name == _DATABASE_URL_ENV_KEY:
        resolved = _resolve_database_url_from_secrets(secrets_mapping)
        if resolved is not None:
            value, source_path = resolved
            _LOGGER.info(
                "config.secrets origin=streamlit action=lookup key=%s source_path=%s status=found",
                name,
                source_path,
            )
            return value
        _LOGGER.info(
            "config.secrets origin=streamlit action=lookup key=%s status=absent",
            name,
        )
        return None

    _LOGGER.info(
        "config.secrets origin=streamlit action=lookup key=%s status=absent",
        name,
    )
    return None


def resolve_dotenv_path(explicit_path: Path | None = None) -> Path:
    if explicit_path is not None:
        return explicit_path.expanduser().resolve()
    return project_root() / ".env"


def load_project_dotenv(explicit_path: Path | None = None) -> DotenvLoadResult:
    """
    Carrega variaveis do .env para os.environ.

    Nao sobrescreve variaveis ja definidas no ambiente (override=False).
    Se o arquivo nao existir ou python-dotenv nao estiver instalado, retorna sem erro.
    """
    env_path = resolve_dotenv_path(explicit_path)

    try:
        from dotenv import load_dotenv
    except ImportError:
        return DotenvLoadResult(
            found=env_path.is_file(),
            loaded=False,
            path=env_path,
            message="Pacote python-dotenv nao instalado; usando apenas variaveis do sistema.",
        )

    if not env_path.is_file():
        return DotenvLoadResult(
            found=False,
            loaded=False,
            path=env_path,
            message="Arquivo .env nao localizado.",
        )

    load_dotenv(env_path, override=False)
    return DotenvLoadResult(
        found=True,
        loaded=True,
        path=env_path,
        message="Arquivo .env carregado com sucesso.",
    )
