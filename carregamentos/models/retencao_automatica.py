from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from carregamentos.models.execucao_retencao import ResultadoRetencao
from carregamentos.models.retencao import GestaoDadosPainel


@dataclass(frozen=True)
class ResultadoRetencaoAutomatica:
    """Resultado da rotina automatica de retenção na inicialização."""

    executado: bool
    mensagem: str
    pacotes_analisados: int
    pacotes_removidos: int
    pacotes_mantidos: int
    espaco_recuperado_bytes: int
    duracao_ms: float
    executado_em: datetime
    resultado_retencao: ResultadoRetencao | None = None
    painel_atualizado: GestaoDadosPainel | None = None
