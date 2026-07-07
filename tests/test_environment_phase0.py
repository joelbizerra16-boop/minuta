from __future__ import annotations

import os
import sqlite3
import tempfile
from pathlib import Path
from unittest.mock import patch

from core.bootstrap import configure_application_storage
from core.environment_diagnostics import (
    DiagnosticStatus,
    run_environment_diagnostics,
    validate_postgresql_connection,
    validate_sqlite_source_url,
)
from core.env_loader import load_project_dotenv
from core.settings import get_settings, reset_settings_cache
from core.startup_environment import bootstrap_environment_from_dotenv, run_startup_environment_checks
from infrastructure.database import get_engine


def _reset_dotenv_state() -> None:
    import core.env_loader as env_loader_module
    import core.settings as settings_module
    import core.startup_environment as startup_module

    settings_module._DOTENV_LOADED = False
    settings_module._DOTENV_RESULT = None
    startup_module._DOTENV_BOOTSTRAPPED = False
    env_loader_module  # keep reference
    reset_settings_cache()


def test_dotenv_loads_variables_from_file() -> None:
    _reset_dotenv_state()
    with tempfile.TemporaryDirectory() as tmp_dir:
        env_path = Path(tmp_dir) / ".env"
        env_path.write_text(
            "MINUTA_DATABASE_URL=sqlite:///" + (Path(tmp_dir) / "from_env.db").as_posix() + "\n"
            "MINUTA_STORAGE_BACKEND=sql\n",
            encoding="utf-8",
        )
        for key in list(os.environ.keys()):
            if key.startswith("MINUTA_"):
                os.environ.pop(key, None)
        result = load_project_dotenv(env_path)
        reset_settings_cache()
        settings = get_settings()
        assert result.loaded is True
        assert "from_env.db" in settings.database_url


def test_missing_dotenv_keeps_sqlite_default() -> None:
    _reset_dotenv_state()
    with tempfile.TemporaryDirectory() as tmp_dir:
        missing = Path(tmp_dir) / "missing.env"
        for key in list(os.environ.keys()):
            if key.startswith("MINUTA_"):
                os.environ.pop(key, None)
        with patch("core.env_loader.resolve_dotenv_path", return_value=missing):
            load_project_dotenv(missing)
            reset_settings_cache()
            settings = get_settings()
        assert settings.database_url.startswith("sqlite:///")
        assert settings.database_url.endswith("minuta.db")


def test_sqlite_source_validation_success() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "origem.db"
        connection = sqlite3.connect(db_path)
        connection.execute("CREATE TABLE demo (id INTEGER PRIMARY KEY)")
        connection.commit()
        connection.close()
        item = validate_sqlite_source_url(f"sqlite:///{db_path.as_posix()}")
        assert item.status == DiagnosticStatus.OK
        assert "localizado com sucesso" in item.message


def test_sqlite_source_validation_missing_file() -> None:
    item = validate_sqlite_source_url("sqlite:///C:/caminho/inexistente/minuta.db")
    assert item.status == DiagnosticStatus.ERROR
    assert "nao encontrado" in item.message.lower()


def test_postgresql_validation_skipped_for_sqlite_url() -> None:
    item = validate_postgresql_connection("sqlite:///data/minuta.db")
    assert item.status == DiagnosticStatus.SKIPPED


def test_startup_environment_checks_never_raise() -> None:
    _reset_dotenv_state()
    report = run_startup_environment_checks()
    assert report.overall_status


def test_bootstrap_still_works_with_sqlite_after_phase0() -> None:
    import infrastructure.database as db_module

    _reset_dotenv_state()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        for key in list(os.environ.keys()):
            if key.startswith("MINUTA_"):
                os.environ.pop(key, None)
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_path / 'boot.db').as_posix()}"
        reset_settings_cache()
        bootstrap_environment_from_dotenv()
        configure_application_storage()
        engine = get_engine()
        assert engine is not None
        engine.dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


def test_diagnostic_report_ready_for_migration_when_fully_configured() -> None:
    _reset_dotenv_state()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        sqlite_path = tmp_path / "origem.db"
        sqlite3.connect(sqlite_path).close()
        env_path = tmp_path / ".env"
        env_path.write_text(
            "\n".join(
                [
                    f"MINUTA_SQLITE_SOURCE_URL=sqlite:///{sqlite_path.as_posix()}",
                    "MINUTA_DATABASE_URL=postgresql+psycopg2://user:pass@localhost:5432/minuta?sslmode=require",
                ]
            ),
            encoding="utf-8",
        )
        with patch("core.environment_diagnostics.validate_postgresql_connection") as mock_pg:
            from core.environment_diagnostics import DiagnosticItem

            mock_pg.return_value = DiagnosticItem(
                label="PostgreSQL",
                status=DiagnosticStatus.OK,
                message="Conexao PostgreSQL validada.",
            )
            with patch("core.environment_diagnostics._check_alembic_cli") as mock_alembic:
                mock_alembic.return_value = DiagnosticItem(
                    label="Alembic",
                    status=DiagnosticStatus.OK,
                    message="CLI Alembic disponivel no PATH.",
                )
                report = run_environment_diagnostics(dotenv_path=env_path, reload_settings=True)
        assert report.ready_for_migration is True
        assert report.overall_status == "AMBIENTE PRONTO PARA MIGRACAO"
