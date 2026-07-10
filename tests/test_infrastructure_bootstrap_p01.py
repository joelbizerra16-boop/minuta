"""Testes da estabilizacao P0.1 — bootstrap once-per-process."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from core.bootstrap import configure_application_storage
from core.environment_diagnostics import get_postgresql_validation_count, run_environment_diagnostics
from core.infrastructure_bootstrap import is_app_bootstrapped, reset_infrastructure_bootstrap_state
from core.settings import get_settings
from core.startup_environment import run_startup_environment_checks
from infrastructure.database import get_engine, get_engine_create_count, is_database_initialized
from infrastructure.schema import get_alembic_run_count


@pytest.fixture(autouse=True)
def _clean_bootstrap_state():
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    yield
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def _dispose_before_tempdir_cleanup() -> None:
    reset_infrastructure_bootstrap_state()


def test_configure_application_storage_runs_once_per_process() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "once.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        configure_application_storage()
        first_engine = get_engine()
        assert is_app_bootstrapped()
        assert get_engine_create_count() == 1

        configure_application_storage()
        configure_application_storage()

        assert get_engine() is first_engine
        assert get_engine_create_count() == 1
        assert is_database_initialized()
        _dispose_before_tempdir_cleanup()


def test_startup_environment_checks_cached_per_process() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "env.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        first = run_startup_environment_checks()
        second = run_startup_environment_checks()
        assert second is first


def test_ensure_full_schema_skips_after_first_verification() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "schema_once.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        configure_application_storage()
        assert get_alembic_run_count() == 0

        configure_application_storage()
        configure_application_storage()
        assert get_alembic_run_count() == 0
        _dispose_before_tempdir_cleanup()


def test_postgresql_validation_cached() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "pg_cache.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        run_environment_diagnostics(reload_settings=False)
        assert get_postgresql_validation_count() == 0

        run_environment_diagnostics(reload_settings=False)
        assert get_postgresql_validation_count() == 0


def test_alembic_not_called_on_subsequent_bootstraps() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "alembic_once.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        with patch("auth.migration.alembic_runner.run_alembic_cli_upgrade") as alembic_mock:
            configure_application_storage()
            configure_application_storage()
            configure_application_storage()
            alembic_mock.assert_not_called()
        _dispose_before_tempdir_cleanup()


def test_services_not_recreated_on_rerun() -> None:
    from carregamentos.bootstrap import get_analise_operacional_service, get_carregamento_repository

    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "services.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        configure_application_storage()
        repo_first = get_carregamento_repository()
        analise_first = get_analise_operacional_service()

        configure_application_storage()

        assert get_carregamento_repository() is repo_first
        assert get_analise_operacional_service() is analise_first
        _dispose_before_tempdir_cleanup()
