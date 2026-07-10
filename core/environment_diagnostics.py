"""Diagnostico e validacao do ambiente (FASE 0 — sem migracao nem alteracao de banco)."""

from __future__ import annotations

import logging
import os
import shutil
import sqlite3
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from sqlalchemy import create_engine, text
from sqlalchemy.engine import make_url

from core.env_loader import DotenvLoadResult, load_project_dotenv, resolve_dotenv_path
from core.settings import AppSettings, get_settings, reset_settings_cache
from infrastructure.persistence.sql_compat import normalize_dialect

_LOGGER = logging.getLogger("minuta.environment")

_POSTGRES_VALIDATION_CACHE: dict[str, DiagnosticItem] = {}
_POSTGRES_VALIDATION_COUNT = 0


def get_postgresql_validation_count() -> int:
    return _POSTGRES_VALIDATION_COUNT


def reset_environment_diagnostics_cache() -> None:
    """Utilitario de teste: limpa cache de validacao PostgreSQL."""
    global _POSTGRES_VALIDATION_CACHE, _POSTGRES_VALIDATION_COUNT
    _POSTGRES_VALIDATION_CACHE.clear()
    _POSTGRES_VALIDATION_COUNT = 0


class DiagnosticStatus(str, Enum):
    OK = "ok"
    WARNING = "warning"
    ERROR = "error"
    SKIPPED = "skipped"


@dataclass(frozen=True)
class DiagnosticItem:
    label: str
    status: DiagnosticStatus
    message: str


@dataclass
class EnvironmentDiagnosticReport:
    items: list[DiagnosticItem] = field(default_factory=list)
    storage_backend: str = ""
    postgres_driver: str = ""
    overall_status: str = ""
    ready_for_migration: bool = False

    def add(self, label: str, status: DiagnosticStatus, message: str) -> None:
        self.items.append(DiagnosticItem(label=label, status=status, message=message))

    def format_banner(self) -> str:
        lines = [
            "",
            "=" * 53,
            "DIAGNOSTICO DO AMBIENTE",
            "=" * 53,
        ]
        for item in self.items:
            symbol = {
                DiagnosticStatus.OK: "[OK]",
                DiagnosticStatus.WARNING: "[AVISO]",
                DiagnosticStatus.ERROR: "[ERRO]",
                DiagnosticStatus.SKIPPED: "[--]",
            }[item.status]
            lines.append(f"{item.label}")
            lines.append(f"{symbol} {item.message}")
            lines.append("")
        if self.storage_backend:
            lines.append(f"Storage Backend")
            lines.append(self.storage_backend)
            lines.append("")
        if self.postgres_driver:
            lines.append(f"Driver PostgreSQL")
            lines.append(self.postgres_driver)
            lines.append("")
        lines.append("Status")
        lines.append(self.overall_status)
        lines.append("=" * 53)
        return "\n".join(lines)


def _friendly_database_error(exc: Exception) -> str:
    text_error = str(exc).strip()
    lowered = text_error.lower()
    if "ssl" in lowered or "sslmode" in lowered:
        return f"Erro SSL na conexao PostgreSQL: {text_error}"
    if "password authentication failed" in lowered or "authentication failed" in lowered:
        return "Erro de autenticacao PostgreSQL: usuario ou senha invalidos."
    if "could not translate host name" in lowered or "name or service not known" in lowered:
        return f"Host PostgreSQL nao encontrado: {text_error}"
    if "connection refused" in lowered:
        return "Conexao PostgreSQL recusada: verifique host, porta e firewall."
    if "timeout" in lowered or "timed out" in lowered:
        return f"Tempo esgotado ao conectar ao PostgreSQL: {text_error}"
    if "no module named 'psycopg2'" in lowered or "psycopg2" in lowered and "import" in lowered:
        return "Driver PostgreSQL ausente (psycopg2)."
    return f"Nao foi possivel conectar ao PostgreSQL: {text_error or type(exc).__name__}"


def _sqlite_path_from_url(url: str) -> Path | None:
    parsed = make_url(url)
    if normalize_dialect(parsed.drivername or "") != "sqlite":
        return None
    database = str(parsed.database or "").strip()
    if not database or database == ":memory:":
        return None
    path = Path(database).expanduser()
    if not path.is_absolute():
        path = path.resolve()
    return path


def validate_sqlite_source_url(sqlite_source_url: str | None) -> DiagnosticItem:
    if not sqlite_source_url:
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.SKIPPED,
            message="MINUTA_SQLITE_SOURCE_URL nao definida.",
        )

    db_path = _sqlite_path_from_url(sqlite_source_url)
    if db_path is None:
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.ERROR,
            message="MINUTA_SQLITE_SOURCE_URL invalida: apenas caminhos de arquivo SQLite sao suportados.",
        )

    if not db_path.exists():
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.ERROR,
            message=f"Banco SQLite nao encontrado: {db_path}",
        )

    if not db_path.is_file():
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.ERROR,
            message=f"Caminho SQLite nao e um arquivo: {db_path}",
        )

    if not os.access(db_path, os.R_OK):
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.ERROR,
            message=f"Sem permissao de leitura no banco SQLite: {db_path}",
        )

    try:
        connection = sqlite3.connect(f"file:{db_path.as_posix()}?mode=ro", uri=True)
        connection.execute("SELECT 1")
        connection.close()
    except sqlite3.Error as exc:
        return DiagnosticItem(
            label="SQLite (origem migracao)",
            status=DiagnosticStatus.ERROR,
            message=f"Banco SQLite localizado, mas nao foi possivel abrir: {exc}",
        )

    return DiagnosticItem(
        label="SQLite (origem migracao)",
        status=DiagnosticStatus.OK,
        message=f"Banco SQLite localizado com sucesso: {db_path}",
    )


def validate_active_sqlite_database(settings: AppSettings) -> DiagnosticItem:
    if normalize_dialect(make_url(settings.database_url).drivername or "") != "sqlite":
        return DiagnosticItem(
            label="SQLite (banco ativo)",
            status=DiagnosticStatus.SKIPPED,
            message="Banco ativo configurado como PostgreSQL.",
        )

    db_path = _sqlite_path_from_url(settings.database_url)
    if db_path is None:
        return DiagnosticItem(
            label="SQLite (banco ativo)",
            status=DiagnosticStatus.OK,
            message="Banco SQLite em memoria ou caminho relativo aceito.",
        )

    if db_path.exists() and db_path.is_file() and os.access(db_path, os.R_OK):
        return DiagnosticItem(
            label="SQLite (banco ativo)",
            status=DiagnosticStatus.OK,
            message=f"Banco SQLite localizado com sucesso: {db_path}",
        )

    return DiagnosticItem(
        label="SQLite (banco ativo)",
        status=DiagnosticStatus.WARNING,
        message=(
            f"Arquivo SQLite ainda nao existe (sera criado na primeira execucao): {db_path}"
            if db_path
            else "Caminho SQLite ativo nao resolvido."
        ),
    )


def validate_postgresql_connection(database_url: str) -> DiagnosticItem:
    global _POSTGRES_VALIDATION_COUNT

    normalized_url = make_url(database_url).render_as_string(hide_password=False)
    cached = _POSTGRES_VALIDATION_CACHE.get(normalized_url)
    if cached is not None:
        _LOGGER.debug("environment.postgresql_validation skipped url=%s", normalized_url)
        return cached

    driver = normalize_dialect(make_url(database_url).drivername or "")
    if driver != "postgresql":
        item = DiagnosticItem(
            label="PostgreSQL",
            status=DiagnosticStatus.SKIPPED,
            message="MINUTA_DATABASE_URL nao aponta para PostgreSQL.",
        )
        _POSTGRES_VALIDATION_CACHE[normalized_url] = item
        return item

    try:
        import psycopg2  # noqa: F401
    except ImportError:
        item = DiagnosticItem(
            label="PostgreSQL",
            status=DiagnosticStatus.ERROR,
            message="Driver PostgreSQL ausente (psycopg2).",
        )
        _POSTGRES_VALIDATION_CACHE[normalized_url] = item
        return item

    _POSTGRES_VALIDATION_COUNT += 1

    try:
        from infrastructure.database import get_configured_database_url, get_engine, is_database_initialized

        if is_database_initialized() and get_configured_database_url() == normalized_url:
            with get_engine().connect() as connection:
                connection.execute(text("SELECT 1"))
            item = DiagnosticItem(
                label="PostgreSQL",
                status=DiagnosticStatus.OK,
                message="Conexao PostgreSQL validada.",
            )
            _POSTGRES_VALIDATION_CACHE[normalized_url] = item
            return item
    except Exception:
        pass

    engine = create_engine(database_url, future=True, pool_pre_ping=True)
    try:
        with engine.connect() as connection:
            connection.execute(text("SELECT 1"))
    except Exception as exc:
        item = DiagnosticItem(
            label="PostgreSQL",
            status=DiagnosticStatus.ERROR,
            message=_friendly_database_error(exc),
        )
        _POSTGRES_VALIDATION_CACHE[normalized_url] = item
        return item
    finally:
        engine.dispose()

    item = DiagnosticItem(
        label="PostgreSQL",
        status=DiagnosticStatus.OK,
        message="Conexao PostgreSQL validada.",
    )
    _POSTGRES_VALIDATION_CACHE[normalized_url] = item
    return item


def _check_alembic_cli() -> DiagnosticItem:
    if shutil.which("alembic"):
        return DiagnosticItem(
            label="Alembic",
            status=DiagnosticStatus.OK,
            message="CLI Alembic disponivel no PATH.",
        )
    return DiagnosticItem(
        label="Alembic",
        status=DiagnosticStatus.WARNING,
        message="CLI Alembic nao encontrada no PATH (necessaria para migrations na FASE 1).",
    )


def _resolve_overall_status(
    *,
    settings: AppSettings,
    dotenv: DotenvLoadResult,
    sqlite_source: DiagnosticItem,
    postgres: DiagnosticItem,
    alembic: DiagnosticItem,
) -> tuple[str, bool]:
    pg_configured = normalize_dialect(make_url(settings.database_url).drivername or "") == "postgresql"
    sqlite_source_configured = bool(settings.sqlite_source_url)

    migration_ready = (
        dotenv.found
        and sqlite_source.status == DiagnosticStatus.OK
        and postgres.status == DiagnosticStatus.OK
        and alembic.status == DiagnosticStatus.OK
    )

    if migration_ready:
        return "AMBIENTE PRONTO PARA MIGRACAO", True

    if pg_configured and postgres.status == DiagnosticStatus.OK and not sqlite_source_configured:
        return "POSTGRESQL CONECTADO — configure MINUTA_SQLITE_SOURCE_URL para migracao", False

    if pg_configured and postgres.status == DiagnosticStatus.ERROR:
        return "POSTGRESQL CONFIGURADO — conexao pendente", False

    if sqlite_source_configured and sqlite_source.status == DiagnosticStatus.ERROR:
        return "SQLITE DE ORIGEM CONFIGURADO — arquivo nao localizado", False

    if not dotenv.found:
        return "SISTEMA OPERACIONAL (SQLite padrao — .env ausente)", False

    if not pg_configured:
        return "SISTEMA OPERACIONAL (SQLite)", False

    return "AMBIENTE PARCIALMENTE CONFIGURADO", False


def run_environment_diagnostics(
    *,
    dotenv_path: Path | None = None,
    reload_settings: bool = True,
) -> EnvironmentDiagnosticReport:
    dotenv = load_project_dotenv(dotenv_path)
    if reload_settings:
        reset_settings_cache()
    settings = get_settings()

    report = EnvironmentDiagnosticReport()
    report.add(
        "Arquivo .env",
        DiagnosticStatus.OK if dotenv.found and dotenv.loaded else DiagnosticStatus.WARNING,
        dotenv.message,
    )

    sqlite_source = validate_sqlite_source_url(settings.sqlite_source_url)
    report.items.append(sqlite_source)

    active_sqlite = validate_active_sqlite_database(settings)
    report.items.append(active_sqlite)

    postgres = validate_postgresql_connection(settings.database_url)
    report.items.append(postgres)

    alembic = _check_alembic_cli()
    report.items.append(alembic)

    report.storage_backend = settings.storage_backend.value.upper()
    driver = make_url(settings.database_url).drivername or ""
    if "postgres" in driver:
        report.postgres_driver = driver
    elif shutil.which("alembic"):
        report.postgres_driver = "psycopg2 (quando MINUTA_DATABASE_URL for PostgreSQL)"
    else:
        report.postgres_driver = "psycopg2"

    overall, ready = _resolve_overall_status(
        settings=settings,
        dotenv=dotenv,
        sqlite_source=sqlite_source,
        postgres=postgres,
        alembic=alembic,
    )
    report.overall_status = overall
    report.ready_for_migration = ready
    return report


def log_environment_diagnostics(report: EnvironmentDiagnosticReport | None = None) -> EnvironmentDiagnosticReport:
    diagnostic = report or run_environment_diagnostics()
    for line in diagnostic.format_banner().splitlines():
        if not line.strip():
            continue
        if line.startswith("✗") or "ERRO" in diagnostic.overall_status.upper() and line == diagnostic.overall_status:
            _LOGGER.warning(line)
        else:
            _LOGGER.info(line)
    return diagnostic
