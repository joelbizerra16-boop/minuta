from __future__ import annotations

import logging
import time
from datetime import date, datetime, timezone

from carregamentos.models.retencao import GestaoDadosPainel
from carregamentos.models.retencao_automatica import ResultadoRetencaoAutomatica
from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService, RetencaoExecucaoError
from carregamentos.services.gestao_dados_service import GestaoDadosService
from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService
from infrastructure.database import get_engine
from infrastructure.models.constants import AUDIT_CATEGORIA_SISTEMA, AUDIT_EVENTO_RETENCAO_DADOS
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.services.database_usage_service import DatabaseUsageService
from infrastructure.unit_of_work import UnitOfWork

_LOGGER = logging.getLogger("minuta.retencao.automatica")


class RetencaoAutomaticaService:
    """Rotina automatica de retenção na inicialização — reutiliza servicos existentes."""

    def __init__(
        self,
        gestao_dados_service: GestaoDadosService | None = None,
        simulacao_service: SimulacaoRetencaoService | None = None,
        execucao_service: ExecucaoRetencaoService | None = None,
        database_usage_service: DatabaseUsageService | None = None,
    ) -> None:
        if gestao_dados_service is None:
            from carregamentos.bootstrap import get_gestao_dados_service

            gestao_dados_service = get_gestao_dados_service()
        if simulacao_service is None:
            from carregamentos.bootstrap import get_simulacao_retencao_service

            simulacao_service = get_simulacao_retencao_service()
        if execucao_service is None:
            from carregamentos.bootstrap import get_execucao_retencao_service

            execucao_service = get_execucao_retencao_service()

        self._gestao = gestao_dados_service
        self._simulacao = simulacao_service
        self._execucao = execucao_service
        self._database_usage = database_usage_service or DatabaseUsageService()

    def executar(self, referencia: date | None = None) -> ResultadoRetencaoAutomatica:
        inicio = time.perf_counter()
        executado_em = datetime.now(timezone.utc)

        try:
            self._validar_conexao_banco()
        except Exception as exc:
            return self._resultado_falha_inicial(
                inicio=inicio,
                executado_em=executado_em,
                mensagem="Banco indisponivel na inicializacao.",
                exc=exc,
                pacotes_analisados=0,
            )

        if not self._gestao.possui_carregamentos_elegiveis(referencia):
            duracao_ms = (time.perf_counter() - inicio) * 1000.0
            mensagem = "Nenhum pacote elegivel encontrado."
            _LOGGER.info(
                "retencao_automatica.concluida resultado=sem_elegiveis duracao_ms=%.2f",
                duracao_ms,
            )
            painel = self._atualizar_indicadores(referencia)
            return ResultadoRetencaoAutomatica(
                executado=False,
                mensagem=mensagem,
                pacotes_analisados=0,
                pacotes_removidos=0,
                pacotes_mantidos=0,
                espaco_recuperado_bytes=0,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                painel_atualizado=painel,
            )

        try:
            relatorio = self._simulacao.executar_simulacao(referencia)
        except Exception as exc:
            return self._resultado_falha_inicial(
                inicio=inicio,
                executado_em=executado_em,
                mensagem="Falha ao simular pacotes elegiveis.",
                exc=exc,
                pacotes_analisados=0,
            )

        pacotes_analisados = len(relatorio.pacotes)
        aptos = relatorio.pacotes_apto_futura_retencao
        mantidos = pacotes_analisados - len(aptos)

        if not aptos:
            duracao_ms = (time.perf_counter() - inicio) * 1000.0
            mensagem = (
                f"Nenhum pacote apto para retencao. "
                f"Analisados={pacotes_analisados}, mantidos={pacotes_analisados}."
            )
            _LOGGER.info(
                "retencao_automatica.concluida resultado=sem_aptos analisados=%s mantidos=%s duracao_ms=%.2f",
                pacotes_analisados,
                pacotes_analisados,
                duracao_ms,
            )
            painel = self._atualizar_indicadores(referencia)
            return ResultadoRetencaoAutomatica(
                executado=False,
                mensagem=mensagem,
                pacotes_analisados=pacotes_analisados,
                pacotes_removidos=0,
                pacotes_mantidos=pacotes_analisados,
                espaco_recuperado_bytes=0,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                painel_atualizado=painel,
            )

        usuario_id, usuario_nome = self._resolver_usuario_sistema()

        try:
            _, confirmacao = self._execucao.preparar_execucao(referencia)
        except RetencaoExecucaoError as exc:
            return self._resultado_falha_inicial(
                inicio=inicio,
                executado_em=executado_em,
                mensagem=str(exc),
                exc=exc,
                pacotes_analisados=pacotes_analisados,
                pacotes_mantidos=mantidos,
            )

        resultado = self._execucao.executar_retencao(
            confirmacao,
            usuario_id=int(usuario_id) if usuario_id is not None else 0,
            usuario_nome=usuario_nome,
            referencia=referencia,
            origem="automatica",
        )
        duracao_ms = (time.perf_counter() - inicio) * 1000.0
        painel = self._atualizar_indicadores(referencia)

        if resultado.sucesso:
            removidos = resultado.carregamentos_removidos
            _LOGGER.info(
                "retencao_automatica.concluida resultado=sucesso analisados=%s removidos=%s "
                "mantidos=%s espaco_bytes=%s duracao_ms=%.2f retencao_ms=%.2f",
                pacotes_analisados,
                removidos,
                mantidos + (pacotes_analisados - removidos - mantidos),
                resultado.espaco_recuperado_bytes,
                duracao_ms,
                resultado.duracao_ms,
            )
            return ResultadoRetencaoAutomatica(
                executado=True,
                mensagem=resultado.mensagem,
                pacotes_analisados=pacotes_analisados,
                pacotes_removidos=removidos,
                pacotes_mantidos=pacotes_analisados - removidos,
                espaco_recuperado_bytes=resultado.espaco_recuperado_bytes,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                resultado_retencao=resultado,
                painel_atualizado=painel,
            )

        self._registrar_auditoria_falha(
            usuario_id=usuario_id,
            mensagem=resultado.mensagem,
            pacotes_analisados=pacotes_analisados,
            pacotes_mantidos=pacotes_analisados,
        )
        _LOGGER.warning(
            "retencao_automatica.concluida resultado=falha mensagem=%s revertido=%s duracao_ms=%.2f",
            resultado.mensagem,
            resultado.revertido,
            duracao_ms,
        )
        return ResultadoRetencaoAutomatica(
            executado=False,
            mensagem=resultado.mensagem,
            pacotes_analisados=pacotes_analisados,
            pacotes_removidos=0,
            pacotes_mantidos=pacotes_analisados,
            espaco_recuperado_bytes=0,
            duracao_ms=duracao_ms,
            executado_em=executado_em,
            resultado_retencao=resultado,
            painel_atualizado=painel,
        )

    @staticmethod
    def _validar_conexao_banco() -> None:
        engine = get_engine()
        with engine.connect() as connection:
            connection.exec_driver_sql("SELECT 1")

    def _atualizar_indicadores(self, referencia: date | None) -> GestaoDadosPainel | None:
        try:
            _ = self._database_usage.medir()
            return self._gestao.obter_painel(referencia)
        except Exception:
            _LOGGER.exception("retencao_automatica.falha_atualizar_indicadores")
            return None

    @staticmethod
    def _resolver_usuario_sistema() -> tuple[int | None, str]:
        try:
            from auth.bootstrap import DEFAULT_ADMIN_USERNAME, get_usuario_repository

            admin = get_usuario_repository().get_by_username(DEFAULT_ADMIN_USERNAME)
            if admin is not None:
                return int(admin.id), "Sistema"
        except Exception:
            _LOGGER.debug("retencao_automatica.usuario_sistema_indisponivel", exc_info=True)
        return None, "Sistema"

    def _resultado_falha_inicial(
        self,
        *,
        inicio: float,
        executado_em: datetime,
        mensagem: str,
        exc: BaseException,
        pacotes_analisados: int,
        pacotes_mantidos: int = 0,
    ) -> ResultadoRetencaoAutomatica:
        duracao_ms = (time.perf_counter() - inicio) * 1000.0
        _LOGGER.exception(
            "retencao_automatica.falha mensagem=%s duracao_ms=%.2f",
            mensagem,
            duracao_ms,
        )
        usuario_id, _ = self._resolver_usuario_sistema()
        self._registrar_auditoria_falha(
            usuario_id=usuario_id,
            mensagem=mensagem,
            pacotes_analisados=pacotes_analisados,
            pacotes_mantidos=pacotes_mantidos or pacotes_analisados,
            detalhe=str(exc),
        )
        painel = self._atualizar_indicadores(None)
        return ResultadoRetencaoAutomatica(
            executado=False,
            mensagem=mensagem,
            pacotes_analisados=pacotes_analisados,
            pacotes_removidos=0,
            pacotes_mantidos=pacotes_mantidos or pacotes_analisados,
            espaco_recuperado_bytes=0,
            duracao_ms=duracao_ms,
            executado_em=executado_em,
            painel_atualizado=painel,
        )

    @staticmethod
    def _registrar_auditoria_falha(
        *,
        usuario_id: int | None,
        mensagem: str,
        pacotes_analisados: int,
        pacotes_mantidos: int,
        detalhe: str | None = None,
    ) -> None:
        try:
            with UnitOfWork() as uow:
                SqlEventoAuditoriaRepository(uow.session).append(
                    EventoAuditoriaRecord(
                        id=0,
                        categoria=AUDIT_CATEGORIA_SISTEMA,
                        evento=AUDIT_EVENTO_RETENCAO_DADOS,
                        usuario_id=usuario_id,
                        entidade_tipo="retencao_automatica",
                        entidade_id=None,
                        descricao=f"Retencao automatica na inicializacao: {mensagem}",
                        metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                            pacotes_analisados=pacotes_analisados,
                            pacotes_mantidos=pacotes_mantidos,
                            resultado="FALHA",
                            origem="automatica",
                            detalhe=detalhe or mensagem,
                        ),
                        ip_origem=None,
                    )
                )
        except Exception:
            _LOGGER.exception("retencao_automatica.falha_auditoria_tecnica")
