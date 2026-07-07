from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from carregamentos.models.retencao_automatica import ResultadoRetencaoAutomatica

_LOGGER = logging.getLogger("minuta.startup.retention")

_RETENCAO_AUTOMATICA_EXECUTADA = False


def reset_startup_retention_flag() -> None:
    """Utilitario de teste para reexecutar a rotina de inicializacao."""
    global _RETENCAO_AUTOMATICA_EXECUTADA
    _RETENCAO_AUTOMATICA_EXECUTADA = False


def run_startup_retention_once() -> ResultadoRetencaoAutomatica | None:
    """Executa a retencao automatica uma unica vez por processo de inicializacao."""
    global _RETENCAO_AUTOMATICA_EXECUTADA
    if _RETENCAO_AUTOMATICA_EXECUTADA:
        return None

    _RETENCAO_AUTOMATICA_EXECUTADA = True

    try:
        from carregamentos.services.retencao_automatica_service import RetencaoAutomaticaService

        resultado = RetencaoAutomaticaService().executar()
        _LOGGER.info(
            "startup.retencao_automatica finalizada executado=%s mensagem=%s",
            resultado.executado,
            resultado.mensagem,
        )
        return resultado
    except Exception:
        _LOGGER.exception("startup.retencao_automatica erro_nao_bloqueante")
        return None
