"""Inicializacao do ambiente na subida da aplicacao (FASE 0)."""

from __future__ import annotations

import logging

from core.environment_diagnostics import (
    EnvironmentDiagnosticReport,
    log_environment_diagnostics,
    run_environment_diagnostics,
)
from core.env_loader import hydrate_runtime_secrets, load_project_dotenv
from core.settings import reset_settings_cache

_LOGGER = logging.getLogger("minuta.startup.environment")

_DOTENV_BOOTSTRAPPED = False
_ENVIRONMENT_VALIDATED = False
_CACHED_ENVIRONMENT_REPORT: EnvironmentDiagnosticReport | None = None


def reset_startup_environment_cache() -> None:
    """Utilitario de teste: permite reexecutar diagnostico de ambiente."""
    global _ENVIRONMENT_VALIDATED, _CACHED_ENVIRONMENT_REPORT
    _ENVIRONMENT_VALIDATED = False
    _CACHED_ENVIRONMENT_REPORT = None


def bootstrap_environment_from_dotenv() -> None:
    """Carrega .env uma vez por processo antes de resolver configuracoes."""
    global _DOTENV_BOOTSTRAPPED
    if _DOTENV_BOOTSTRAPPED:
        return
    hydrate_runtime_secrets()
    result = load_project_dotenv()
    reset_settings_cache()
    _DOTENV_BOOTSTRAPPED = True
    if result.loaded:
        _LOGGER.info(result.message)
    else:
        _LOGGER.info(result.message)


def run_startup_environment_checks() -> EnvironmentDiagnosticReport:
    """
    Executa diagnostico amigavel do ambiente.

    Nao executa migrations, ETL nem altera banco.
    Nunca propaga excecoes para a interface do usuario.
    """
    global _ENVIRONMENT_VALIDATED, _CACHED_ENVIRONMENT_REPORT

    bootstrap_environment_from_dotenv()
    if _ENVIRONMENT_VALIDATED and _CACHED_ENVIRONMENT_REPORT is not None:
        _LOGGER.debug("startup.environment_checks skipped reutilizando diagnostico em cache")
        return _CACHED_ENVIRONMENT_REPORT

    try:
        report = log_environment_diagnostics(run_environment_diagnostics(reload_settings=False))
    except Exception:
        _LOGGER.exception("Falha ao executar diagnostico do ambiente; inicializacao continua.")
        report = EnvironmentDiagnosticReport()
        report.overall_status = "DIAGNOSTICO INDISPONIVEL — inicializacao continua em modo padrao"

    _ENVIRONMENT_VALIDATED = True
    _CACHED_ENVIRONMENT_REPORT = report
    return report
