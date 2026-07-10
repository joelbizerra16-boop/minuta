"""Carregamento do arquivo .env e secrets do runtime (Streamlit Cloud)."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

_STREAMLIT_PROMOTED_KEYS: set[str] = set()

# Caminhos aninhados comuns em secrets.toml do Streamlit Cloud.
_NESTED_DATABASE_URL_PATHS: tuple[tuple[str, ...], ...] = (
    ("MINUTA_DATABASE_URL",),
    ("minuta", "database_url"),
    ("neon", "database_url"),
    ("database", "url"),
    ("connections", "postgresql", "url"),
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


def _promote_env_var(name: str, value: str) -> None:
    if os.getenv(name):
        return
    os.environ[name] = value
    _STREAMLIT_PROMOTED_KEYS.add(name)


def _promote_streamlit_scalar_secrets(store: Mapping[str, object]) -> None:
    for key, value in store.items():
        if _is_scalar_secret(value):
            _promote_env_var(str(key), str(value))


def _promote_database_url_from_nested_secrets(store: Mapping[str, object]) -> None:
    if os.getenv("MINUTA_DATABASE_URL"):
        return
    for path in _NESTED_DATABASE_URL_PATHS:
        resolved = _lookup_nested_secret(store, path)
        if resolved:
            _promote_env_var("MINUTA_DATABASE_URL", resolved)
            return


def hydrate_runtime_secrets() -> bool:
    """
    Hidrata os.environ a partir de st.secrets antes da resolucao de configuracao.

    No Streamlit Cloud, MINUTA_DATABASE_URL costuma existir apenas em secrets.toml.
    O acesso a st.secrets promove chaves escalares de topo para os.environ; esta
    funcao garante essa hidratacao antes do bootstrap e cobre layouts aninhados.
    """
    try:
        import streamlit as st
    except ImportError:
        return False

    try:
        secrets_store = st.secrets
        if hasattr(secrets_store, "to_dict"):
            secrets_mapping = secrets_store.to_dict()
        else:
            secrets_mapping = dict(secrets_store)
    except Exception:
        return False

    if not secrets_mapping:
        return False

    _promote_streamlit_scalar_secrets(secrets_mapping)
    _promote_database_url_from_nested_secrets(secrets_mapping)
    return bool(_STREAMLIT_PROMOTED_KEYS)


def lookup_runtime_secret(name: str) -> str | None:
    """Leitura direta de st.secrets para chaves escalares (fallback de get_env)."""
    try:
        import streamlit as st
    except ImportError:
        return None

    try:
        value = st.secrets[name]
    except Exception:
        return None

    if not _is_scalar_secret(value):
        return None
    stripped = str(value).strip()
    return stripped or None


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
