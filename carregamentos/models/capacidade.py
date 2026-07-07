from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from enum import Enum

from infrastructure.services.database_usage_service import UsoBancoDados


class FaixaCapacidade(str, Enum):
    VERDE = "VERDE"
    AMARELA = "AMARELA"
    LARANJA = "LARANJA"
    VERMELHA = "VERMELHA"

    @property
    def rotulo_banner(self) -> str:
        return {
            FaixaCapacidade.VERDE: "Normal",
            FaixaCapacidade.AMARELA: "Atencao",
            FaixaCapacidade.LARANJA: "Critica",
            FaixaCapacidade.VERMELHA: "Critica",
        }[self]

    @property
    def cor_hex(self) -> str:
        return {
            FaixaCapacidade.VERDE: "#16A34A",
            FaixaCapacidade.AMARELA: "#CA8A04",
            FaixaCapacidade.LARANJA: "#EA580C",
            FaixaCapacidade.VERMELHA: "#DC2626",
        }[self]


@dataclass(frozen=True)
class CapacidadeOperacional:
    uso_banco: UsoBancoDados
    faixa: FaixaCapacidade
    barra_visual: str
    exibir_aviso_discreto: bool
    exibir_alerta_vermelho: bool
    requer_dialogo_login: bool

    @property
    def percentual(self) -> float | None:
        return self.uso_banco.utilizacao_percentual


@dataclass(frozen=True)
class PreviaRetencaoCapacidade:
    data_alvo: date
    carregamentos: int
    notas_fiscais: int
    documentos_pdf: int
    documentos_xml: int
    eventos: int
    historicos: int
    espaco_recuperavel_bytes: int
    espaco_atual_bytes: int | None
    espaco_apos_bytes: int | None
    percentual_atual: float | None
    percentual_apos: float | None
    carregamento_ids: tuple[int, ...]
