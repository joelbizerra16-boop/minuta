from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from carregamentos.models.capacidade import CapacidadeOperacional
from infrastructure.services.database_usage_service import UsoBancoDados


@dataclass(frozen=True)
class RetencaoContagensArvore:
    """Contagens agregadas da arvore operacional abaixo de carregamentos elegiveis."""

    carregamentos_elegiveis: int
    notas_fiscais: int
    itens_carregamento: int
    itens_nota_fiscal: int
    documentos_xml: int
    documentos_pdf: int
    historicos: int
    eventos: int
    caminhos_pdf: tuple[str, ...]
    espaco_xmls_bytes: int


@dataclass(frozen=True)
class PacoteRetencao:
    """
    Pacote de Retencao — representa a arvore operacional elegivel (carregamento como raiz).

    Conceito interno para futura execucao de retencao (Fase R2). Nao persiste no banco.
    """

    data_corte: date
    carregamentos: int
    itens_carregamento: int
    notas_fiscais: int
    itens_nota_fiscal: int
    documentos_xml: int
    documentos_pdf: int
    historicos: int
    eventos: int
    espaco_pdfs_bytes: int
    espaco_xmls_bytes: int
    espaco_metadados_sql_bytes: int

    @property
    def espaco_recuperavel_bytes(self) -> int:
        return int(self.espaco_pdfs_bytes) + int(self.espaco_xmls_bytes) + int(self.espaco_metadados_sql_bytes)

    @property
    def possui_elegiveis(self) -> bool:
        return int(self.carregamentos) > 0

    @classmethod
    def from_contagens(
        cls,
        *,
        data_corte: date,
        contagens: RetencaoContagensArvore,
        espaco_pdfs_bytes: int,
        espaco_metadados_sql_bytes: int,
    ) -> PacoteRetencao:
        return cls(
            data_corte=data_corte,
            carregamentos=contagens.carregamentos_elegiveis,
            itens_carregamento=contagens.itens_carregamento,
            notas_fiscais=contagens.notas_fiscais,
            itens_nota_fiscal=contagens.itens_nota_fiscal,
            documentos_xml=contagens.documentos_xml,
            documentos_pdf=contagens.documentos_pdf,
            historicos=contagens.historicos,
            eventos=contagens.eventos,
            espaco_pdfs_bytes=espaco_pdfs_bytes,
            espaco_xmls_bytes=contagens.espaco_xmls_bytes,
            espaco_metadados_sql_bytes=espaco_metadados_sql_bytes,
        )


@dataclass(frozen=True)
class RetencaoPreview:
    """Resultado legado da analise de retencao (somente leitura)."""

    periodo_inicio: date
    periodo_fim: date
    dias_mantidos: int
    data_corte: date
    pacote: PacoteRetencao
    simulacao: bool = True

    @property
    def carregamentos_elegiveis(self) -> int:
        return self.pacote.carregamentos

    @property
    def notas_fiscais(self) -> int:
        return self.pacote.notas_fiscais

    @property
    def itens_carregamento(self) -> int:
        return self.pacote.itens_carregamento

    @property
    def itens_nota_fiscal(self) -> int:
        return self.pacote.itens_nota_fiscal

    @property
    def documentos_xml(self) -> int:
        return self.pacote.documentos_xml

    @property
    def documentos_pdf(self) -> int:
        return self.pacote.documentos_pdf

    @property
    def historicos(self) -> int:
        return self.pacote.historicos

    @property
    def eventos(self) -> int:
        return self.pacote.eventos

    @property
    def espaco_pdfs_bytes(self) -> int:
        return self.pacote.espaco_pdfs_bytes

    @property
    def espaco_xmls_bytes(self) -> int:
        return self.pacote.espaco_xmls_bytes

    @property
    def espaco_db_estimado_bytes(self) -> int:
        return self.pacote.espaco_metadados_sql_bytes

    @property
    def espaco_total_estimado_bytes(self) -> int:
        return self.pacote.espaco_recuperavel_bytes

    @property
    def possui_elegiveis(self) -> bool:
        return self.pacote.possui_elegiveis


@dataclass(frozen=True)
class GestaoDadosPainel:
    """Dashboard administrativo de gestao de dados (analise — sem exclusao)."""

    uso_banco: UsoBancoDados
    capacidade: CapacidadeOperacional
    politica_dias_mantidos: int
    politica_descricao: str
    politica_status: str
    periodo_inicio: date
    periodo_fim: date
    data_corte: date
    pacote: PacoteRetencao
    simulacao: bool = True

    @property
    def possui_elegiveis(self) -> bool:
        return self.pacote.possui_elegiveis
