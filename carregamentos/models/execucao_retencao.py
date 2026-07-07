from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime


@dataclass(frozen=True)
class ConfirmacaoRetencao:
    """Resumo exibido ao administrador antes da exclusao irreversivel."""

    carregamentos: int
    notas_fiscais: int
    documentos_xml: int
    documentos_pdf: int
    eventos: int
    historicos: int
    espaco_estimado_bytes: int
    data_corte: date
    carregamento_ids: tuple[int, ...]

    @property
    def possui_pacotes(self) -> bool:
        return self.carregamentos > 0


@dataclass(frozen=True)
class ResultadoRetencao:
    sucesso: bool
    mensagem: str
    carregamentos_removidos: int
    notas_fiscais_removidas: int
    documentos_xml_removidos: int
    documentos_pdf_removidos: int
    eventos_removidos: int
    historicos_removidos: int
    espaco_recuperado_bytes: int
    duracao_ms: float
    executado_em: datetime
    arquivos_pdf_removidos: int
    arquivos_xml_removidos: int
    arquivos_falha: tuple[str, ...] = ()
    revertido: bool = False
