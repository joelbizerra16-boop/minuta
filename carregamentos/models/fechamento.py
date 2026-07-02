from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from carregamentos.models.carregamento import Carregamento, NfHistoricoConflito

FechamentoStatus = Literal[
    "primeira_impressao",
    "reimpressao",
    "complementacao",
    "needs_reentrega",
    "needs_reimpressao_confirm",
    "invalid",
    "error",
]


@dataclass(frozen=True)
class ImpressaoInfo:
    carregamento_id: int
    numero_carregamento: str
    primeira_impressao_data: str
    primeira_impressao_usuario: str
    quantidade_impressoes: int


@dataclass(frozen=True)
class FechamentoResult:
    status: FechamentoStatus
    carregamento: Carregamento | None = None
    conflitos: tuple[NfHistoricoConflito, ...] = ()
    impressao_info: ImpressaoInfo | None = None
    message: str = ""
