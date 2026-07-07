from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
from typing import TYPE_CHECKING

from carregamentos.models.retencao import GestaoDadosPainel, PacoteRetencao, RetencaoPreview
from carregamentos.repository.retencao_repository import RetencaoRepository
from carregamentos.repository.sql_retencao_repository import SqlRetencaoRepository
from core.retention_policy import (
    RETENTION_DAYS,
    RETENTION_POLICY_STATUS,
    retention_days_before_today,
    retention_policy_description,
)
from infrastructure.database import get_pdf_storage_dir
from infrastructure.services.database_usage_service import DatabaseUsageService

if TYPE_CHECKING:
    from carregamentos.services.gestao_capacidade_service import GestaoCapacidadeService

_BYTES_ESTIMADOS_POR_REGISTRO = {
    "carregamento": 520,
    "item_carregamento": 220,
    "documento": 320,
    "historico": 280,
    "evento": 420,
    "item_nota_fiscal": 160,
}
_BYTES_PDF_FALLBACK = 70_000


class GestaoDadosService:
    """Gestao de dados do sistema: analise, retencao e uso do banco (sem exclusao)."""

    def __init__(
        self,
        repository: RetencaoRepository | None = None,
        pdf_storage_dir: Path | None = None,
        database_usage_service: DatabaseUsageService | None = None,
        capacidade_service: GestaoCapacidadeService | None = None,
    ) -> None:
        self._repository = repository or SqlRetencaoRepository()
        self._pdf_storage_dir = pdf_storage_dir
        self._database_usage_service = database_usage_service or DatabaseUsageService()
        if capacidade_service is None:
            from carregamentos.services.gestao_capacidade_service import GestaoCapacidadeService

            capacidade_service = GestaoCapacidadeService(gestao_dados_service=self)
        self._capacidade_service = capacidade_service

    @staticmethod
    def calcular_data_corte(referencia: date | None = None) -> date:
        base = referencia or date.today()
        return base - timedelta(days=retention_days_before_today())

    @staticmethod
    def obter_periodo_mantido(referencia: date | None = None) -> tuple[date, date]:
        fim = referencia or date.today()
        inicio = fim - timedelta(days=retention_days_before_today())
        return inicio, fim

    def possui_carregamentos_elegiveis(self, referencia: date | None = None) -> bool:
        return self._repository.possui_carregamentos_elegiveis(self.calcular_data_corte(referencia))

    def montar_pacote_retencao(self, referencia: date | None = None) -> PacoteRetencao:
        data_corte = self.calcular_data_corte(referencia)
        contagens = self._repository.coletar_contagens_arvore(data_corte)
        espaco_pdfs_bytes = self._estimar_espaco_pdfs(contagens.caminhos_pdf)
        espaco_metadados_sql_bytes = self._estimar_espaco_metadados_sql(contagens)
        return PacoteRetencao.from_contagens(
            data_corte=data_corte,
            contagens=contagens,
            espaco_pdfs_bytes=espaco_pdfs_bytes,
            espaco_metadados_sql_bytes=espaco_metadados_sql_bytes,
        )

    def analisar(self, referencia: date | None = None) -> RetencaoPreview:
        periodo_inicio, periodo_fim = self.obter_periodo_mantido(referencia)
        pacote = self.montar_pacote_retencao(referencia)
        return RetencaoPreview(
            periodo_inicio=periodo_inicio,
            periodo_fim=periodo_fim,
            dias_mantidos=RETENTION_DAYS,
            data_corte=pacote.data_corte,
            pacote=pacote,
            simulacao=True,
        )

    def obter_painel(self, referencia: date | None = None) -> GestaoDadosPainel:
        periodo_inicio, periodo_fim = self.obter_periodo_mantido(referencia)
        pacote = self.montar_pacote_retencao(referencia)
        uso_banco = self._database_usage_service.medir()
        capacidade = self._capacidade_service.avaliar_capacidade(uso_banco)
        return GestaoDadosPainel(
            uso_banco=uso_banco,
            capacidade=capacidade,
            politica_dias_mantidos=RETENTION_DAYS,
            politica_descricao=retention_policy_description(),
            politica_status=RETENTION_POLICY_STATUS,
            periodo_inicio=periodo_inicio,
            periodo_fim=periodo_fim,
            data_corte=pacote.data_corte,
            pacote=pacote,
            simulacao=True,
        )

    def _estimar_espaco_pdfs(self, caminhos: tuple[str, ...]) -> int:
        if not caminhos:
            return 0
        base_dir = self._pdf_storage_dir or get_pdf_storage_dir()
        total = 0
        for relative_path in caminhos:
            candidate = Path(relative_path)
            absolute = candidate if candidate.is_absolute() else base_dir / relative_path
            if absolute.is_file():
                total += int(absolute.stat().st_size)
            else:
                total += _BYTES_PDF_FALLBACK
        return total

    @staticmethod
    def _estimar_espaco_metadados_sql(contagens) -> int:
        return (
            contagens.carregamentos_elegiveis * _BYTES_ESTIMADOS_POR_REGISTRO["carregamento"]
            + contagens.itens_carregamento * _BYTES_ESTIMADOS_POR_REGISTRO["item_carregamento"]
            + contagens.documentos_pdf * _BYTES_ESTIMADOS_POR_REGISTRO["documento"]
            + contagens.historicos * _BYTES_ESTIMADOS_POR_REGISTRO["historico"]
            + contagens.eventos * _BYTES_ESTIMADOS_POR_REGISTRO["evento"]
            + contagens.itens_nota_fiscal * _BYTES_ESTIMADOS_POR_REGISTRO["item_nota_fiscal"]
        )


GestaoRetencaoService = GestaoDadosService
