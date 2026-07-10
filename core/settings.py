from __future__ import annotations

import os
from dataclasses import dataclass
from enum import Enum
from functools import lru_cache
from pathlib import Path

from sqlalchemy.engine import make_url

from core.database_config import (
    DatabaseUrlSource,
    RuntimeEnvironment,
    enforce_production_database_policy,
    reset_database_config_state,
    resolve_database_url,
    resolve_runtime_environment,
    snapshot_environment_before_dotenv,
)
from core.env_loader import DotenvLoadResult, load_project_dotenv, resolve_dotenv_path
from infrastructure.persistence.sql_compat import normalize_dialect

_DOTENV_LOADED = False
_DOTENV_RESULT: DotenvLoadResult | None = None


class StorageBackend(str, Enum):
    SQL = "sql"
    JSON = "json"
    DUAL = "dual"


def _default_database_url() -> str:
    data_root = (Path.cwd() / "data").resolve()
    db_path = data_root / "minuta.db"
    return f"sqlite:///{db_path.as_posix()}"


def _default_sqlite_source_url() -> str | None:
    default_db = Path(_default_database_url().replace("sqlite:///", "", 1))
    if default_db.is_file():
        return f"sqlite:///{default_db.as_posix()}"
    return None


def ensure_dotenv_loaded() -> DotenvLoadResult:
    """Carrega .env uma vez por processo antes de ler variaveis."""
    global _DOTENV_LOADED, _DOTENV_RESULT
    if not _DOTENV_LOADED:
        snapshot_environment_before_dotenv()
        _DOTENV_RESULT = load_project_dotenv()
        _DOTENV_LOADED = True
    return _DOTENV_RESULT or load_project_dotenv()


def get_env(name: str, default: str | None = None) -> str | None:
    """Leitura centralizada de variaveis de ambiente (apos .env)."""
    ensure_dotenv_loaded()
    value = os.getenv(name)
    if value is None:
        return default
    stripped = str(value).strip()
    return stripped if stripped else default


def get_env_int(name: str, default: int) -> int:
    raw = get_env(name)
    if raw is None:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def get_env_bool(name: str, default: bool = False) -> bool:
    raw = get_env(name)
    if raw is None:
        return default
    return raw.lower() in {"1", "true", "yes", "on"}


def reset_settings_cache() -> None:
    reset_database_config_state()
    get_settings.cache_clear()


def derive_storage_paths(database_url: str) -> tuple[Path, Path, Path]:
    """Deriva data_root e diretorios de arquivo a partir de MINUTA_DATABASE_URL."""
    url = make_url(database_url)
    driver = normalize_dialect(url.drivername or "")

    if driver == "sqlite":
        database = str(url.database or "").strip() or ":memory:"
        if database == ":memory:":
            data_root = (Path.cwd() / "data").resolve()
        else:
            data_root = Path(database).expanduser().resolve().parent
    else:
        configured_root = get_env("MINUTA_DATA_ROOT")
        if configured_root:
            data_root = Path(configured_root).expanduser().resolve()
        else:
            data_root = (Path.cwd() / "data").resolve()

    return (
        data_root,
        data_root / "documentos",
        data_root / "xml_storage",
    )


@dataclass(frozen=True)
class AppSettings:
    storage_backend: StorageBackend
    database_url: str
    database_url_source: DatabaseUrlSource
    runtime_environment: RuntimeEnvironment
    sqlite_source_url: str | None
    data_root: Path
    pdf_storage_dir: Path
    xml_storage_dir: Path
    echo_sql: bool
    retention_days: int
    database_limit_bytes: int
    migration_batch_size: int
    test_postgres_url: str | None
    env_file_path: Path
    env_file_loaded: bool


@lru_cache(maxsize=1)
def get_settings() -> AppSettings:
    dotenv = ensure_dotenv_loaded()

    raw_backend = str(get_env("MINUTA_STORAGE_BACKEND", StorageBackend.SQL.value) or StorageBackend.SQL.value).lower()
    try:
        storage_backend = StorageBackend(raw_backend)
    except ValueError:
        storage_backend = StorageBackend.SQL

    runtime_environment = resolve_runtime_environment(minuta_env=get_env("MINUTA_ENV"))
    database_resolution = resolve_database_url(
        minuta_database_url=get_env("MINUTA_DATABASE_URL"),
        dotenv_loaded=dotenv.loaded,
        runtime_environment=runtime_environment,
    )
    enforce_production_database_policy(database_resolution)
    database_url = database_resolution.database_url
    sqlite_source_url = get_env("MINUTA_SQLITE_SOURCE_URL") or _default_sqlite_source_url()
    data_root, pdf_storage_dir, xml_storage_dir = derive_storage_paths(database_url)

    return AppSettings(
        storage_backend=storage_backend,
        database_url=database_url,
        database_url_source=database_resolution.source,
        runtime_environment=database_resolution.runtime_environment,
        sqlite_source_url=sqlite_source_url,
        data_root=data_root,
        pdf_storage_dir=pdf_storage_dir,
        xml_storage_dir=xml_storage_dir,
        echo_sql=get_env_bool("MINUTA_SQL_ECHO", default=False),
        retention_days=get_env_int("MINUTA_RETENTION_DAYS", 8),
        database_limit_bytes=get_env_int("MINUTA_DATABASE_LIMIT_BYTES", 500 * 1024 * 1024),
        migration_batch_size=get_env_int("MINUTA_MIGRATION_BATCH_SIZE", 500),
        test_postgres_url=get_env("MINUTA_TEST_POSTGRES_URL"),
        env_file_path=resolve_dotenv_path(),
        env_file_loaded=dotenv.loaded,
    )
