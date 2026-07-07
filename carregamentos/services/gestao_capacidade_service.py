from __future__ import annotations

from datetime import date
from typing import TYPE_CHECKING

from carregamentos.models.capacidade import CapacidadeOperacional, FaixaCapacidade, PreviaRetencaoCapacidade
from carregamentos.models.execucao_retencao import ConfirmacaoRetencao
from carregamentos.models.retencao import PacoteRetencao
from carregamentos.repository.retencao_repository import RetencaoRepository
from carregamentos.repository.sql_retencao_repository import SqlRetencaoRepository
from carregamentos.services.gestao_dados_service import GestaoDadosService
from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService
from core.retention_policy import (
    CAPACITY_ORANGE_MIN_PERCENT,
    CAPACITY_RED_MIN_PERCENT,
    CAPACITY_YELLOW_MIN_PERCENT,
    DATABASE_STORAGE_LIMIT_BYTES,
)
from infrastructure.models.constants import AUDIT_CATEGORIA_SISTEMA, AUDIT_EVENTO_CAPACIDADE_ALERTA
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.services.database_usage_service import DatabaseUsageService, UsoBancoDados
from infrastructure.unit_of_work import UnitOfWork

if TYPE_CHECKING:
    from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService


class GestaoCapacidadeError(Exception):
    """Falha na gestao preventiva de capacidade."""


class GestaoCapacidadeService:
    """Gestao preventiva da capacidade operacional do banco Neon."""

    def __init__(
        self,
        gestao_dados_service: GestaoDadosService | None = None,
        database_usage_service: DatabaseUsageService | None = None,
        retencao_repository: RetencaoRepository | None = None,
        simulacao_service: SimulacaoRetencaoService | None = None,
        execucao_service: ExecucaoRetencaoService | None = None,
    ) -> None:
        self._gestao_dados = gestao_dados_service or GestaoDadosService()
        self._database_usage = database_usage_service or DatabaseUsageService()
        self._repository = retencao_repository or SqlRetencaoRepository()
        self._simulacao = simulacao_service
        self._execucao = execucao_service

    def avaliar_capacidade(self, uso: UsoBancoDados | None = None) -> CapacidadeOperacional:
        medicao = uso or self._database_usage.medir()
        faixa = classificar_faixa_capacidade(medicao.utilizacao_percentual)
        percentual = medicao.utilizacao_percentual or 0.0
        return CapacidadeOperacional(
            uso_banco=medicao,
            faixa=faixa,
            barra_visual=montar_barra_capacidade(percentual),
            exibir_aviso_discreto=faixa == FaixaCapacidade.AMARELA,
            exibir_alerta_vermelho=faixa == FaixaCapacidade.VERMELHA,
            requer_dialogo_login=(medicao.utilizacao_percentual or 0) >= CAPACITY_ORANGE_MIN_PERCENT,
        )

    def montar_pacote_dia_mais_antigo(self, referencia: date | None = None) -> PacoteRetencao | None:
        data_corte = self._gestao_dados.calcular_data_corte(referencia)
        data_alvo = self._repository.obter_data_mais_antiga_elegivel(data_corte)
        if data_alvo is None:
            return None
        contagens = self._repository.coletar_contagens_arvore_por_data(data_corte, data_alvo)
        if contagens.carregamentos_elegiveis == 0:
            return None
        espaco_pdfs = self._gestao_dados._estimar_espaco_pdfs(contagens.caminhos_pdf)
        espaco_sql = self._gestao_dados._estimar_espaco_metadados_sql(contagens)
        return PacoteRetencao.from_contagens(
            data_corte=data_corte,
            contagens=contagens,
            espaco_pdfs_bytes=espaco_pdfs,
            espaco_metadados_sql_bytes=espaco_sql,
        )

    def montar_previa_dia_mais_antigo(self, referencia: date | None = None) -> PreviaRetencaoCapacidade | None:
        data_corte = self._gestao_dados.calcular_data_corte(referencia)
        data_alvo = self._repository.obter_data_mais_antiga_elegivel(data_corte)
        if data_alvo is None:
            return None

        carregamento_ids = self._repository.listar_carregamento_ids_por_data(data_corte, data_alvo)
        if not carregamento_ids:
            return None

        pacote = self.montar_pacote_dia_mais_antigo(referencia)
        if pacote is None:
            return None

        uso = self._database_usage.medir()
        recuperavel = pacote.espaco_recuperavel_bytes
        ocupado_apos, pct_apos = projetar_uso_apos_recuperacao(uso, recuperavel)

        return PreviaRetencaoCapacidade(
            data_alvo=data_alvo,
            carregamentos=pacote.carregamentos,
            notas_fiscais=pacote.notas_fiscais,
            documentos_pdf=pacote.documentos_pdf,
            documentos_xml=pacote.documentos_xml,
            eventos=pacote.eventos,
            historicos=pacote.historicos,
            espaco_recuperavel_bytes=recuperavel,
            espaco_atual_bytes=uso.bytes_ocupados,
            espaco_apos_bytes=ocupado_apos,
            percentual_atual=uso.utilizacao_percentual,
            percentual_apos=pct_apos,
            carregamento_ids=carregamento_ids,
        )

    def preparar_retencao_dia_mais_antigo(
        self,
        referencia: date | None = None,
    ) -> tuple[PreviaRetencaoCapacidade, ConfirmacaoRetencao]:
        previa = self.montar_previa_dia_mais_antigo(referencia)
        if previa is None:
            raise GestaoCapacidadeError("Nao ha dia elegivel para retencao sugerida.")

        execucao = self._resolver_execucao()
        relatorio, confirmacao = execucao.preparar_execucao_por_carregamentos(
            previa.carregamento_ids,
            referencia=referencia,
        )
        _ = relatorio
        return previa, confirmacao

    def registrar_auditoria_capacidade(
        self,
        *,
        usuario_id: int | None,
        capacidade: CapacidadeOperacional,
        previa: PreviaRetencaoCapacidade | None = None,
        ip_origem: str | None = None,
    ) -> None:
        with UnitOfWork() as uow:
            repo = SqlEventoAuditoriaRepository(uow.session)
            repo.append(
                EventoAuditoriaRecord(
                    id=0,
                    categoria=AUDIT_CATEGORIA_SISTEMA,
                    evento=AUDIT_EVENTO_CAPACIDADE_ALERTA,
                    usuario_id=usuario_id,
                    entidade_tipo="capacidade_operacional",
                    entidade_id=None,
                    descricao="Alerta preventivo de capacidade do banco de dados.",
                    metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                        percentual=capacidade.percentual,
                        bytes_ocupados=capacidade.uso_banco.bytes_ocupados,
                        bytes_limite=capacidade.uso_banco.bytes_limite,
                        faixa=capacidade.faixa.value,
                        espaco_recuperavel_previsto=(
                            previa.espaco_recuperavel_bytes if previa is not None else None
                        ),
                        data_sugerida=str(previa.data_alvo) if previa is not None else None,
                    ),
                    ip_origem=ip_origem,
                )
            )

    def _resolver_execucao(self) -> ExecucaoRetencaoService:
        from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService

        if self._execucao is not None:
            return self._execucao
        if self._simulacao is None:
            return ExecucaoRetencaoService()
        return ExecucaoRetencaoService(simulacao_service=self._simulacao)


def classificar_faixa_capacidade(percentual: float | None) -> FaixaCapacidade:
    if percentual is None:
        return FaixaCapacidade.VERDE
    if percentual >= CAPACITY_RED_MIN_PERCENT:
        return FaixaCapacidade.VERMELHA
    if percentual >= CAPACITY_ORANGE_MIN_PERCENT:
        return FaixaCapacidade.LARANJA
    if percentual >= CAPACITY_YELLOW_MIN_PERCENT:
        return FaixaCapacidade.AMARELA
    return FaixaCapacidade.VERDE


def montar_barra_capacidade(percentual: float) -> str:
    pct = max(min(float(percentual), 100.0), 0.0)
    preenchido = int(round(pct / 10.0))
    if pct > 0 and preenchido == 0:
        preenchido = 1
    preenchido = min(preenchido, 10)
    return ("█" * preenchido) + ("░" * (10 - preenchido))


def projetar_uso_apos_recuperacao(
    uso: UsoBancoDados,
    bytes_recuperaveis: int,
) -> tuple[int | None, float | None]:
    if uso.bytes_ocupados is None or uso.bytes_limite is None or uso.bytes_limite <= 0:
        return None, None
    ocupado_apos = max(int(uso.bytes_ocupados) - max(int(bytes_recuperaveis), 0), 0)
    pct_apos = round((ocupado_apos / int(uso.bytes_limite)) * 100, 1)
    return ocupado_apos, pct_apos


def limite_operacional_bytes() -> int:
    return DATABASE_STORAGE_LIMIT_BYTES
