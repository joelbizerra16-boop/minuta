from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date

from sqlalchemy.orm import Session

from carregamentos.repository.simulacao_retencao_repository import ArvoreCarregamentoRaw


@dataclass(frozen=True)
class RecursosCompartilhadosRemovidos:
    nota_fiscal_ids: tuple[int, ...]
    chaves_xml: tuple[str, ...]
    caminhos_xml: tuple[str, ...]


class ExecucaoRetencaoRepository(ABC):
    @abstractmethod
    def carregar_arvores_por_ids(self, session: Session, carregamento_ids: list[int]) -> list[ArvoreCarregamentoRaw]:
        raise NotImplementedError

    @abstractmethod
    def validar_carregamento_elegivel(
        self,
        session: Session,
        carregamento_id: int,
        data_corte: date,
    ) -> bool:
        raise NotImplementedError

    @abstractmethod
    def excluir_arvore_carregamento(self, session: Session, arvore: ArvoreCarregamentoRaw) -> None:
        raise NotImplementedError

    @abstractmethod
    def excluir_recursos_compartilhados_orfos(
        self,
        session: Session,
        candidatos_nf_ids: set[int],
        candidatos_chaves_xml: set[str],
    ) -> RecursosCompartilhadosRemovidos:
        raise NotImplementedError
