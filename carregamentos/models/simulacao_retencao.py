from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum


class SaudePacote(str, Enum):
    SAUDAVEL = "SAUDAVEL"
    ATENCAO = "ATENCAO"
    CRITICO = "CRITICO"

    @property
    def rotulo(self) -> str:
        return {
            SaudePacote.SAUDAVEL: "Saudavel",
            SaudePacote.ATENCAO: "Atencao",
            SaudePacote.CRITICO: "Critico",
        }[self]

    @property
    def emoji(self) -> str:
        return {
            SaudePacote.SAUDAVEL: "🟢",
            SaudePacote.ATENCAO: "🟡",
            SaudePacote.CRITICO: "🔴",
        }[self]


@dataclass(frozen=True)
class CarregamentoElegivelRef:
    id: int
    numero_carregamento: str
    data: date


@dataclass(frozen=True)
class DocumentoPdfValidacao:
    documento_id: int
    tipo: str
    caminho_arquivo: str
    nome_arquivo: str
    hash_sha256: str
    existe_arquivo: bool
    tamanho_bytes: int


@dataclass(frozen=True)
class DocumentoXmlValidacao:
    chave_nfe: str
    numero_nf: str
    documento_xml_id: int | None
    caminho_arquivo: str | None
    hash_sha256: str | None
    registro_ativo: bool
    existe_arquivo: bool
    tamanho_bytes: int


@dataclass(frozen=True)
class ProblemaIntegridade:
    severidade: SaudePacote
    categoria: str
    descricao: str
    carregamento_id: int | None = None
    referencia: str | None = None


@dataclass(frozen=True)
class PacoteRetencaoUnitario:
    """Pacote de Retencao completo para um unico carregamento elegivel."""

    carregamento_id: int
    numero_carregamento: str
    data_carregamento: date
    itens_carregamento: int
    notas_fiscais: int
    itens_nota_fiscal: int
    documentos_xml: int
    documentos_pdf: int
    historicos: int
    eventos: int
    arquivos_encontrados: int
    arquivos_ausentes: int
    integridade_percentual: float
    saude: SaudePacote
    apto_retencao: bool
    problemas: tuple[str, ...] = ()
    pdfs: tuple[DocumentoPdfValidacao, ...] = ()
    xmls: tuple[DocumentoXmlValidacao, ...] = ()
    espaco_pdfs_bytes: int = 0
    espaco_xmls_bytes: int = 0
    espaco_metadados_sql_bytes: int = 0

    @property
    def espaco_recuperavel_bytes(self) -> int:
        return int(self.espaco_pdfs_bytes) + int(self.espaco_xmls_bytes) + int(self.espaco_metadados_sql_bytes)

    @property
    def total_registros(self) -> int:
        return (
            1
            + self.itens_carregamento
            + self.notas_fiscais
            + self.itens_nota_fiscal
            + self.documentos_xml
            + self.documentos_pdf
            + self.historicos
            + self.eventos
        )


@dataclass(frozen=True)
class ResumoSaudeSimulacao:
    pacotes_elegiveis: int
    pacotes_integros: int
    pacotes_com_inconsistencia: int
    integridade_geral_percentual: float
    todos_pdfs_encontrados: bool
    todos_xmls_encontrados: bool
    pacotes_saudaveis: int
    pacotes_atencao: int
    pacotes_criticos: int


@dataclass(frozen=True)
class RelatorioSimulacaoRetencao:
    data_corte: date
    pacotes: tuple[PacoteRetencaoUnitario, ...]
    orfaos: tuple[ProblemaIntegridade, ...]
    resumo: ResumoSaudeSimulacao
    registros_analisados: int
    arquivos_pdf: int
    arquivos_xml: int
    espaco_elegivel_bytes: int
    duracao_ms: float
    simulacao: bool = True

    @property
    def pacotes_apto_futura_retencao(self) -> tuple[PacoteRetencaoUnitario, ...]:
        return tuple(p for p in self.pacotes if p.apto_retencao)

    @property
    def pacotes_requerem_correcao(self) -> tuple[PacoteRetencaoUnitario, ...]:
        return tuple(p for p in self.pacotes if not p.apto_retencao)
