"""Inicializacao do ambiente na subida da aplicacao (FASE 0)."""

from __future__ import annotations

import logging

from core.environment_diagnostics import (
    EnvironmentDiagnosticReport,
    log_environment_diagnostics,
    run_environment_diagnostics,
)
from core.env_loader import load_project_dotenv
from core.settings import reset_settings_cache

_LOGGER = logging.getLogger("minuta.startup.environment")

_DOTENV_BOOTSTRAPPED = False


def bootstrap_environment_from_dotenv() -> None:
    """Carrega .env uma vez por processo antes de resolver configuracoes."""
    global _DOTENV_BOOTSTRAPPED
    if _DOTENV_BOOTSTRAPPED:
        return
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
    bootstrap_environment_from_dotenv()
    try:
        return log_environment_diagnostics(run_environment_diagnostics(reload_settings=False))
    except Exception:
        _LOGGER.exception("Falha ao executar diagnostico do ambiente; inicializacao continua.")
        report = EnvironmentDiagnosticReport()
        report.overall_status = "DIAGNOSTICO INDISPONIVEL — inicializacao continua em modo padrao"
        return report
