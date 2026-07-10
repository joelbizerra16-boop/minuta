"""Estado e metricas do bootstrap de infraestrutura (once-per-process)."""

from __future__ import annotations

_APP_BOOTSTRAPPED = False
_BOOTSTRAPPED_DATABASE_URL: str | None = None


def is_app_bootstrapped() -> bool:
    return _APP_BOOTSTRAPPED


def get_bootstrapped_database_url() -> str | None:
    return _BOOTSTRAPPED_DATABASE_URL


def mark_app_bootstrapped(*, database_url: str) -> None:
    global _APP_BOOTSTRAPPED, _BOOTSTRAPPED_DATABASE_URL
    _APP_BOOTSTRAPPED = True
    _BOOTSTRAPPED_DATABASE_URL = database_url


def reset_app_bootstrap_state() -> None:
    global _APP_BOOTSTRAPPED, _BOOTSTRAPPED_DATABASE_URL
    _APP_BOOTSTRAPPED = False
    _BOOTSTRAPPED_DATABASE_URL = None


def reset_infrastructure_bootstrap_state() -> None:
    """Utilitario de teste: limpa todo estado once-per-process da infraestrutura."""
    from auth import bootstrap as auth_bootstrap
    from carregamentos import bootstrap as carregamentos_bootstrap
    from core import startup_environment
    from core.environment_diagnostics import reset_environment_diagnostics_cache
    from infrastructure.database import reset_database_state
    from infrastructure.schema import reset_schema_bootstrap_state

    reset_app_bootstrap_state()
    reset_environment_diagnostics_cache()
    startup_environment.reset_startup_environment_cache()
    from core.runtime_data_coherence import invalidate_data_signature_cache

    invalidate_data_signature_cache()
    reset_database_state()
    reset_schema_bootstrap_state()
    auth_bootstrap.reset_auth_bootstrap_state()
    carregamentos_bootstrap.reset_carregamentos_bootstrap_state()
