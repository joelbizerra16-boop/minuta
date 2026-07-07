from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date

from carregamentos.models.simulacao_retencao import CarregamentoElegivelRef, ProblemaIntegridade


@dataclass(frozen=True)
class ArvoreCarregamentoRaw:
    carregamento_id: int
    numero_carregamento: str
    data: date
    item_ids: tuple[int, ...]
    chaves_nfe: tuple[str, ...]
    numeros_nf: tuple[str, ...]
    nota_fiscal_ids: tuple[int, ...]
    documento_ids: tuple[int, ...]
    documento_tipos: tuple[str, ...]
    documento_caminhos: tuple[str, ...]
    documento_nomes: tuple[str, ...]
    documento_hashes: tuple[str, ...]
    historico_ids: tuple[int, ...]
    evento_ids: tuple[int, ...]
    item_nota_fiscal_ids: tuple[int, ...]


@dataclass(frozen=True)
class DocumentoXmlRaw:
    id: int
    chave_nfe: str
    numero_nf: str
    caminho_arquivo: str
    hash_sha256: str
    tamanho: int
    ativo: bool


class SimulacaoRetencaoRepository(ABC):
    @abstractmethod
    def listar_carregamentos_elegiveis(self, data_corte: date) -> list[CarregamentoElegivelRef]:
        raise NotImplementedError

    @abstractmethod
    def carregar_arvores_elegiveis(self, data_corte: date) -> list[ArvoreCarregamentoRaw]:
        raise NotImplementedError

    @abstractmethod
    def carregar_documentos_xml_por_chaves(self, chaves: list[str]) -> dict[str, DocumentoXmlRaw]:
        raise NotImplementedError

    @abstractmethod
    def detectar_orfaos(self, data_corte: date, carregamento_ids: list[int]) -> list[ProblemaIntegridade]:
        raise NotImplementedError
