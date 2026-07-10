from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from core.database_config import (
    DatabaseUrlSource,
    ProductionDatabaseConfigurationError,
    RuntimeEnvironment,
    enforce_production_database_policy,
    resolve_database_url,
    resolve_runtime_environment,
)
from core.env_loader import hydrate_runtime_secrets, reset_streamlit_secret_state
from core.settings import get_settings, reset_settings_cache


def test_resolve_runtime_environment_explicit_production() -> None:
    assert resolve_runtime_environment(minuta_env="production") == RuntimeEnvironment.PRODUCTION


def test_resolve_runtime_environment_explicit_development() -> None:
    assert resolve_runtime_environment(minuta_env="development") == RuntimeEnvironment.DEVELOPMENT


def test_resolve_runtime_environment_streamlit_cloud_path() -> None:
    with patch("core.database_config.Path.cwd", return_value=Path("/mount/src/minuta")):
        assert resolve_runtime_environment(minuta_env=None) == RuntimeEnvironment.PRODUCTION


def test_production_without_database_url_raises() -> None:
    with pytest.raises(ProductionDatabaseConfigurationError, match="MINUTA_DATABASE_URL nao configurada"):
        resolve_database_url(
            minuta_database_url=None,
            dotenv_loaded=False,
            runtime_environment=RuntimeEnvironment.PRODUCTION,
        )


def test_production_with_sqlite_url_raises() -> None:
    resolution = resolve_database_url(
        minuta_database_url="sqlite:////tmp/minuta.db",
        dotenv_loaded=True,
        runtime_environment=RuntimeEnvironment.PRODUCTION,
    )
    with pytest.raises(ProductionDatabaseConfigurationError, match="PostgreSQL"):
        enforce_production_database_policy(resolution)


def test_development_fallback_sqlite_when_unconfigured() -> None:
    resolution = resolve_database_url(
        minuta_database_url=None,
        dotenv_loaded=False,
        runtime_environment=RuntimeEnvironment.DEVELOPMENT,
    )
    assert resolution.source == DatabaseUrlSource.SQLITE_DEFAULT
    assert resolution.database_url.startswith("sqlite:///")


def test_get_settings_marks_postgres_from_dotenv() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    with tempfile.TemporaryDirectory() as tmp_dir:
        env_path = Path(tmp_dir) / ".env"
        env_path.write_text(
            "MINUTA_DATABASE_URL=postgresql+psycopg2://user:pass@localhost:5432/minuta?sslmode=require\n",
            encoding="utf-8",
        )
        for key in list(os.environ.keys()):
            if key.startswith("MINUTA_"):
                os.environ.pop(key, None)
        settings_module._DOTENV_LOADED = False
        settings_module._DOTENV_RESULT = None
        with patch("core.env_loader.resolve_dotenv_path", return_value=env_path):
            reset_settings_cache()
            settings = get_settings()
        assert settings.database_url_source == DatabaseUrlSource.DOTENV
        assert settings.runtime_environment == RuntimeEnvironment.DEVELOPMENT
        assert "postgresql" in settings.database_url
    reset_settings_cache()


def test_get_settings_production_mode_requires_database_url() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    with tempfile.TemporaryDirectory() as tmp_dir:
        missing = Path(tmp_dir) / "missing.env"
        for key in list(os.environ.keys()):
            if key.startswith("MINUTA_"):
                os.environ.pop(key, None)
        os.environ["MINUTA_ENV"] = "production"
        settings_module._DOTENV_LOADED = False
        settings_module._DOTENV_RESULT = None
        with patch("core.env_loader.hydrate_runtime_secrets", return_value=False):
            with patch("core.env_loader.resolve_dotenv_path", return_value=missing):
                reset_settings_cache()
                with pytest.raises(ProductionDatabaseConfigurationError):
                    get_settings()
        os.environ.pop("MINUTA_ENV", None)
    reset_settings_cache()


def test_hydrate_streamlit_secrets_promotes_database_url() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    reset_streamlit_secret_state()
    postgres_url = "postgresql+psycopg2://user:pass@neon.example/neondb?sslmode=require"
    for key in list(os.environ.keys()):
        if key.startswith("MINUTA_"):
            os.environ.pop(key, None)

    fake_secrets = MagicMock()
    fake_secrets.to_dict.return_value = {"MINUTA_DATABASE_URL": postgres_url}

    settings_module._DOTENV_LOADED = False
    settings_module._DOTENV_RESULT = None
    mock_streamlit = MagicMock()
    mock_streamlit.secrets = fake_secrets
    with patch.dict(sys.modules, {"streamlit": mock_streamlit}):
        with patch("core.database_config.Path.cwd", return_value=Path("/mount/src/minuta")):
            with patch("core.env_loader.resolve_dotenv_path", return_value=Path("/mount/src/minuta/missing.env")):
                reset_settings_cache()
                settings = get_settings()

    assert settings.runtime_environment == RuntimeEnvironment.PRODUCTION
    assert settings.database_url == postgres_url
    assert settings.database_url_source == DatabaseUrlSource.STREAMLIT_SECRETS
    reset_settings_cache()


def test_hydrate_streamlit_nested_secret_promotes_database_url() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    reset_streamlit_secret_state()
    postgres_url = "postgresql+psycopg2://user:pass@neon.example/neondb?sslmode=require"
    for key in list(os.environ.keys()):
        if key.startswith("MINUTA_"):
            os.environ.pop(key, None)

    fake_secrets = MagicMock()
    fake_secrets.to_dict.return_value = {"minuta": {"database_url": postgres_url}}

    settings_module._DOTENV_LOADED = False
    settings_module._DOTENV_RESULT = None
    mock_streamlit = MagicMock()
    mock_streamlit.secrets = fake_secrets
    with patch.dict(sys.modules, {"streamlit": mock_streamlit}):
        with patch("core.database_config.Path.cwd", return_value=Path("/mount/src/minuta")):
            with patch("core.env_loader.resolve_dotenv_path", return_value=Path("/mount/src/minuta/missing.env")):
                reset_settings_cache()
                settings = get_settings()

    assert settings.database_url == postgres_url
    assert settings.database_url_source == DatabaseUrlSource.STREAMLIT_SECRETS
    reset_settings_cache()


def test_hydrate_streamlit_connections_neon_url() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    reset_streamlit_secret_state()
    postgres_url = "postgresql+psycopg2://user:pass@neon.example/neondb?sslmode=require"
    for key in list(os.environ.keys()):
        if key.startswith("MINUTA_"):
            os.environ.pop(key, None)

    fake_secrets = MagicMock()
    fake_secrets.to_dict.return_value = {
        "connections": {"neon": {"url": postgres_url}},
    }

    settings_module._DOTENV_LOADED = False
    settings_module._DOTENV_RESULT = None
    mock_streamlit = MagicMock()
    mock_streamlit.secrets = fake_secrets
    with patch.dict(sys.modules, {"streamlit": mock_streamlit}):
        with patch("core.database_config.Path.cwd", return_value=Path("/mount/src/minuta")):
            with patch("core.env_loader.resolve_dotenv_path", return_value=Path("/mount/src/minuta/missing.env")):
                reset_settings_cache()
                settings = get_settings()

    assert settings.runtime_environment == RuntimeEnvironment.PRODUCTION
    assert settings.database_url == postgres_url
    assert settings.database_url_source == DatabaseUrlSource.STREAMLIT_SECRETS
    reset_settings_cache()


def test_production_wrong_secret_key_still_raises() -> None:
    import core.settings as settings_module

    reset_settings_cache()
    reset_streamlit_secret_state()
    for key in list(os.environ.keys()):
        if key.startswith("MINUTA_"):
            os.environ.pop(key, None)

    fake_secrets = MagicMock()
    fake_secrets.to_dict.return_value = {
        "DATABASE_URL": "postgresql+psycopg2://user:pass@neon.example/neondb?sslmode=require",
    }

    settings_module._DOTENV_LOADED = False
    settings_module._DOTENV_RESULT = None
    mock_streamlit = MagicMock()
    mock_streamlit.secrets = fake_secrets
    with patch.dict(sys.modules, {"streamlit": mock_streamlit}):
        with patch("core.database_config.Path.cwd", return_value=Path("/mount/src/minuta")):
            with patch("core.env_loader.resolve_dotenv_path", return_value=Path("/mount/src/minuta/missing.env")):
                reset_settings_cache()
                with pytest.raises(ProductionDatabaseConfigurationError):
                    get_settings()
    reset_settings_cache()
