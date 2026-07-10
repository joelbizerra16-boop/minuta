"""Resolucao e politicas de configuracao do banco (camada de bootstrap)."""

from __future__ import annotations

import os
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from sqlalchemy.engine import make_url

_STREAMLIT_CLOUD_ROOT_PREFIX = "/mount/src/"
_ENV_SNAPSHOT_BEFORE_DOTENV: dict[str, str | None] = {}


def _normalize_dialect(driver: str) -> str:
    normalized = str(driver or "").strip().lower()
    if "postgres" in normalized:
        return "postgresql"
    if "sqlite" in normalized:
        return "sqlite"
    return normalized


class RuntimeEnvironment(str, Enum):
    DEVELOPMENT = "development"
    PRODUCTION = "production"


class DatabaseUrlSource(str, Enum):
    ENVIRONMENT = "environment"
    DOTENV = "dotenv"
    SQLITE_DEFAULT = "sqlite_default"


class ProductionDatabaseConfigurationError(RuntimeError):
    """Producao exige PostgreSQL configurado explicitamente via MINUTA_DATABASE_URL."""


@dataclass(frozen=True)
class DatabaseUrlResolution:
    database_url: str
    source: DatabaseUrlSource
    runtime_environment: RuntimeEnvironment


def reset_database_config_state() -> None:
    _ENV_SNAPSHOT_BEFORE_DOTENV.clear()


def snapshot_environment_before_dotenv() -> None:
    """Registra variaveis de ambiente presentes antes do carregamento do .env."""
    _ENV_SNAPSHOT_BEFORE_DOTENV["MINUTA_DATABASE_URL"] = os.getenv("MINUTA_DATABASE_URL")


def _default_database_url() -> str:
    data_root = (Path.cwd() / "data").resolve()
    db_path = data_root / "minuta.db"
    return f"sqlite:///{db_path.as_posix()}"


def _is_streamlit_cloud_runtime() -> bool:
    return str(Path.cwd()).replace("\\", "/").startswith(_STREAMLIT_CLOUD_ROOT_PREFIX)


def resolve_runtime_environment(*, minuta_env: str | None) -> RuntimeEnvironment:
    normalized = str(minuta_env or "").strip().lower()
    if normalized in {"production", "prod"}:
        return RuntimeEnvironment.PRODUCTION
    if normalized in {"development", "dev", "local"}:
        return RuntimeEnvironment.DEVELOPMENT
    if _is_streamlit_cloud_runtime():
        return RuntimeEnvironment.PRODUCTION
    return RuntimeEnvironment.DEVELOPMENT


def _database_url_from_snapshot() -> str | None:
    raw = _ENV_SNAPSHOT_BEFORE_DOTENV.get("MINUTA_DATABASE_URL", os.getenv("MINUTA_DATABASE_URL"))
    if raw is None:
        return None
    stripped = str(raw).strip()
    return stripped or None


def resolve_database_url(
    *,
    minuta_database_url: str | None,
    dotenv_loaded: bool,
    runtime_environment: RuntimeEnvironment,
) -> DatabaseUrlResolution:
    configured = str(minuta_database_url or "").strip()
    if configured:
        if _database_url_from_snapshot():
            source = DatabaseUrlSource.ENVIRONMENT
        elif dotenv_loaded:
            source = DatabaseUrlSource.DOTENV
        else:
            source = DatabaseUrlSource.ENVIRONMENT
        return DatabaseUrlResolution(
            database_url=configured,
            source=source,
            runtime_environment=runtime_environment,
        )

    if runtime_environment == RuntimeEnvironment.PRODUCTION:
        raise ProductionDatabaseConfigurationError(
            "MINUTA_DATABASE_URL nao configurada. Em producao o PostgreSQL (Neon) e obrigatorio. "
            "Configure MINUTA_DATABASE_URL nos Secrets do Streamlit Cloud (ou variaveis de ambiente) "
            "com a URL postgresql+psycopg2://...?sslmode=require."
        )

    return DatabaseUrlResolution(
        database_url=_default_database_url(),
        source=DatabaseUrlSource.SQLITE_DEFAULT,
        runtime_environment=runtime_environment,
    )


def enforce_production_database_policy(resolution: DatabaseUrlResolution) -> None:
    if resolution.runtime_environment != RuntimeEnvironment.PRODUCTION:
        return

    if resolution.source == DatabaseUrlSource.SQLITE_DEFAULT:
        raise ProductionDatabaseConfigurationError(
            "Fallback SQLite nao e permitido em producao. Configure MINUTA_DATABASE_URL para PostgreSQL (Neon)."
        )

    driver = _normalize_dialect(make_url(resolution.database_url).drivername or "")
    if driver != "postgresql":
        raise ProductionDatabaseConfigurationError(
            "MINUTA_DATABASE_URL em producao deve apontar para PostgreSQL (Neon). "
            f"Driver recebido: {driver or 'desconhecido'}."
        )


def describe_database_target(database_url: str) -> dict[str, str]:
    url = make_url(database_url)
    dialect = _normalize_dialect(str(url.drivername or ""))
    return {
        "dialect": dialect,
        "driver": str(url.drivername or ""),
        "host": str(url.host or ""),
        "database": str(url.database or ""),
        "schema": "public" if dialect == "postgresql" else "",
    }
