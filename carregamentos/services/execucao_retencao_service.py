from __future__ import annotations

import time
from datetime import date, datetime, timezone
from pathlib import Path

from carregamentos.models.execucao_retencao import ConfirmacaoRetencao, ResultadoRetencao
from carregamentos.models.simulacao_retencao import PacoteRetencaoUnitario, RelatorioSimulacaoRetencao, SaudePacote
from carregamentos.repository.execucao_retencao_repository import ExecucaoRetencaoRepository
from carregamentos.repository.simulacao_retencao_repository import ArvoreCarregamentoRaw
from carregamentos.repository.sql_execucao_retencao_repository import SqlExecucaoRetencaoRepository
from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService
from infrastructure.models.constants import AUDIT_CATEGORIA_SISTEMA, AUDIT_EVENTO_RETENCAO_DADOS
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.unit_of_work import UnitOfWork


class RetencaoExecucaoError(Exception):
    """Falha na preparacao ou execucao da retencao operacional."""


class ExecucaoRetencaoService:
    """Execucao transacional da retencao manual de pacotes elegiveis e aptos."""

    def __init__(
        self,
        simulacao_service: SimulacaoRetencaoService | None = None,
        repository: ExecucaoRetencaoRepository | None = None,
        pdf_storage_dir: Path | None = None,
        xml_storage_dir: Path | None = None,
    ) -> None:
        self._simulacao = simulacao_service or SimulacaoRetencaoService(
            pdf_storage_dir=pdf_storage_dir,
            xml_storage_dir=xml_storage_dir,
        )
        self._repository = repository or SqlExecucaoRetencaoRepository()
        self._pdf_storage_dir = pdf_storage_dir
        self._xml_storage_dir = xml_storage_dir

    def preparar_execucao(self, referencia: date | None = None) -> tuple[RelatorioSimulacaoRetencao, ConfirmacaoRetencao]:
        relatorio = self._simulacao.executar_simulacao(referencia)
        aptos = relatorio.pacotes_apto_futura_retencao
        if not aptos:
            raise RetencaoExecucaoError("Nenhum pacote elegivel e apto para retencao.")

        for pacote in aptos:
            if pacote.saude == SaudePacote.CRITICO or not pacote.apto_retencao:
                raise RetencaoExecucaoError(
                    f"Pacote {pacote.numero_carregamento} nao atende aos criterios de retencao."
                )

        confirmacao = self._montar_confirmacao(relatorio, aptos)
        return relatorio, confirmacao

    def preparar_execucao_por_carregamentos(
        self,
        carregamento_ids: tuple[int, ...] | list[int],
        referencia: date | None = None,
    ) -> tuple[RelatorioSimulacaoRetencao, ConfirmacaoRetencao]:
        ids_solicitados = tuple(sorted(int(item) for item in carregamento_ids))
        if not ids_solicitados:
            raise RetencaoExecucaoError("Nenhum carregamento informado para retencao.")

        relatorio = self._simulacao.executar_simulacao(referencia)
        aptos_por_id = {
            pacote.carregamento_id: pacote for pacote in relatorio.pacotes_apto_futura_retencao
        }

        aptos_filtrados: list[PacoteRetencaoUnitario] = []
        for carregamento_id in ids_solicitados:
            pacote = aptos_por_id.get(carregamento_id)
            if pacote is None:
                raise RetencaoExecucaoError(
                    f"Carregamento {carregamento_id} nao esta apto para retencao."
                )
            if pacote.saude == SaudePacote.CRITICO or not pacote.apto_retencao:
                raise RetencaoExecucaoError(
                    f"Pacote {pacote.numero_carregamento} nao atende aos criterios de retencao."
                )
            aptos_filtrados.append(pacote)

        confirmacao = self._montar_confirmacao(relatorio, tuple(aptos_filtrados))
        return relatorio, confirmacao

    def executar_retencao(
        self,
        confirmacao: ConfirmacaoRetencao,
        *,
        usuario_id: int,
        usuario_nome: str,
        ip_origem: str | None = None,
        referencia: date | None = None,
        origem: str = "manual",
    ) -> ResultadoRetencao:
        inicio = time.perf_counter()
        executado_em = datetime.now(timezone.utc)

        try:
            relatorio, confirmacao_atual = self.preparar_execucao_por_carregamentos(
                confirmacao.carregamento_ids,
                referencia=referencia,
            )
        except RetencaoExecucaoError as exc:
            duracao_ms = (time.perf_counter() - inicio) * 1000.0
            return ResultadoRetencao(
                sucesso=False,
                mensagem=str(exc),
                carregamentos_removidos=0,
                notas_fiscais_removidas=0,
                documentos_xml_removidos=0,
                documentos_pdf_removidos=0,
                eventos_removidos=0,
                historicos_removidos=0,
                espaco_recuperado_bytes=0,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                arquivos_pdf_removidos=0,
                arquivos_xml_removidos=0,
                revertido=True,
            )

        if confirmacao_atual.carregamento_ids != confirmacao.carregamento_ids:
            duracao_ms = (time.perf_counter() - inicio) * 1000.0
            return ResultadoRetencao(
                sucesso=False,
                mensagem="Os pacotes elegiveis mudaram desde a confirmacao. Execute a simulacao novamente.",
                carregamentos_removidos=0,
                notas_fiscais_removidas=0,
                documentos_xml_removidos=0,
                documentos_pdf_removidos=0,
                eventos_removidos=0,
                historicos_removidos=0,
                espaco_recuperado_bytes=0,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                arquivos_pdf_removidos=0,
                arquivos_xml_removidos=0,
                revertido=True,
            )

        ids_confirmados = set(confirmacao.carregamento_ids)
        aptos = tuple(
            pacote
            for pacote in relatorio.pacotes_apto_futura_retencao
            if pacote.carregamento_id in ids_confirmados
        )
        pacotes_por_id = {pacote.carregamento_id: pacote for pacote in aptos}
        data_corte = relatorio.data_corte

        candidatos_nf: set[int] = set()
        candidatos_chaves: set[str] = set()
        caminhos_pdf: list[str] = []
        totais = {
            "eventos": 0,
            "historicos": 0,
            "pdfs": 0,
            "itens": 0,
        }

        compartilhados_nfs: tuple[int, ...] = ()
        compartilhados_xml_chaves: tuple[str, ...] = ()

        for pacote in aptos:
            for pdf in pacote.pdfs:
                if pdf.caminho_arquivo:
                    caminhos_pdf.append(pdf.caminho_arquivo)
            for xml in pacote.xmls:
                if xml.chave_nfe:
                    candidatos_chaves.add(xml.chave_nfe)

        arquivos_xml_planejados: list[str] = []
        espaco_planejado = sum(pacote.espaco_recuperavel_bytes for pacote in aptos)

        try:
            with UnitOfWork() as uow:
                session = uow.session
                arvores = self._repository.carregar_arvores_por_ids(
                    session,
                    list(confirmacao.carregamento_ids),
                )
                if len(arvores) != len(confirmacao.carregamento_ids):
                    raise RetencaoExecucaoError("Falha ao montar a arvore dos pacotes elegiveis.")

                for arvore in arvores:
                    pacote = pacotes_por_id.get(arvore.carregamento_id)
                    if pacote is None:
                        raise RetencaoExecucaoError(
                            f"Pacote {arvore.numero_carregamento} nao esta apto para retencao."
                        )
                    self._validar_pacote_antes_exclusao(pacote, arvore, data_corte)
                    if not self._repository.validar_carregamento_elegivel(session, arvore.carregamento_id, data_corte):
                        raise RetencaoExecucaoError(
                            f"Carregamento {arvore.numero_carregamento} nao esta mais elegivel."
                        )

                    candidatos_nf.update(int(nf_id) for nf_id in arvore.nota_fiscal_ids)
                    candidatos_chaves.update(str(chave).strip() for chave in arvore.chaves_nfe if str(chave).strip())
                    totais["eventos"] += len(arvore.evento_ids)
                    totais["historicos"] += len(arvore.historico_ids)
                    totais["pdfs"] += len(arvore.documento_ids)
                    totais["itens"] += len(arvore.item_ids)

                    self._repository.excluir_arvore_carregamento(session, arvore)

                compartilhados = self._repository.excluir_recursos_compartilhados_orfos(
                    session,
                    candidatos_nf,
                    candidatos_chaves,
                )
                compartilhados_nfs = compartilhados.nota_fiscal_ids
                compartilhados_xml_chaves = compartilhados.chaves_xml
                arquivos_xml_planejados = list(compartilhados.caminhos_xml)

                audit_repo = SqlEventoAuditoriaRepository(session)
                rotulo_origem = "automatica" if origem == "automatica" else "manual"
                audit_repo.append(
                    EventoAuditoriaRecord(
                        id=0,
                        categoria=AUDIT_CATEGORIA_SISTEMA,
                        evento=AUDIT_EVENTO_RETENCAO_DADOS,
                        usuario_id=int(usuario_id) if int(usuario_id) > 0 else None,
                        entidade_tipo="retencao_operacional",
                        entidade_id=None,
                        descricao=(
                            f"Retencao {rotulo_origem} executada por {usuario_nome}: "
                            f"{len(confirmacao.carregamento_ids)} carregamento(s) removido(s)."
                        ),
                        metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                            carregamentos=len(confirmacao.carregamento_ids),
                            notas_fiscais=len(compartilhados.nota_fiscal_ids),
                            documentos_xml=len(compartilhados.chaves_xml),
                            documentos_pdf=totais["pdfs"],
                            eventos=totais["eventos"],
                            historicos=totais["historicos"],
                            espaco_recuperado_bytes=espaco_planejado,
                            resultado="SUCESSO",
                            origem=rotulo_origem,
                        ),
                        ip_origem=ip_origem,
                    )
                )
        except Exception as exc:
            duracao_ms = (time.perf_counter() - inicio) * 1000.0
            mensagem = str(exc) if isinstance(exc, RetencaoExecucaoError) else "Falha transacional na retencao."
            return ResultadoRetencao(
                sucesso=False,
                mensagem=mensagem,
                carregamentos_removidos=0,
                notas_fiscais_removidas=0,
                documentos_xml_removidos=0,
                documentos_pdf_removidos=0,
                eventos_removidos=0,
                historicos_removidos=0,
                espaco_recuperado_bytes=0,
                duracao_ms=duracao_ms,
                executado_em=executado_em,
                arquivos_pdf_removidos=0,
                arquivos_xml_removidos=0,
                revertido=True,
            )

        pdf_removidos, pdf_falhas = self._remover_arquivos(caminhos_pdf, base_dir=self._resolver_pdf_dir())
        xml_removidos, xml_falhas = self._remover_arquivos(arquivos_xml_planejados, base_dir=self._resolver_xml_dir())
        falhas = tuple(pdf_falhas + xml_falhas)

        try:
            from carregamentos.bootstrap import invalidate_analise_operacional_cache

            invalidate_analise_operacional_cache()
        except Exception:
            pass

        duracao_ms = (time.perf_counter() - inicio) * 1000.0
        return ResultadoRetencao(
            sucesso=True,
            mensagem="Retencao concluida com sucesso.",
            carregamentos_removidos=len(confirmacao.carregamento_ids),
            notas_fiscais_removidas=len(compartilhados_nfs),
            documentos_xml_removidos=len(compartilhados_xml_chaves),
            documentos_pdf_removidos=totais["pdfs"],
            eventos_removidos=totais["eventos"],
            historicos_removidos=totais["historicos"],
            espaco_recuperado_bytes=espaco_planejado,
            duracao_ms=duracao_ms,
            executado_em=executado_em,
            arquivos_pdf_removidos=pdf_removidos,
            arquivos_xml_removidos=xml_removidos,
            arquivos_falha=falhas,
            revertido=False,
        )

    @staticmethod
    def _montar_confirmacao(
        relatorio: RelatorioSimulacaoRetencao,
        aptos: tuple[PacoteRetencaoUnitario, ...],
    ) -> ConfirmacaoRetencao:
        chaves_nf: set[str] = set()
        for pacote in aptos:
            for xml in pacote.xmls:
                if xml.chave_nfe:
                    chaves_nf.add(xml.chave_nfe)

        return ConfirmacaoRetencao(
            carregamentos=len(aptos),
            notas_fiscais=sum(pacote.notas_fiscais for pacote in aptos),
            documentos_xml=len(chaves_nf),
            documentos_pdf=sum(pacote.documentos_pdf for pacote in aptos),
            eventos=sum(pacote.eventos for pacote in aptos),
            historicos=sum(pacote.historicos for pacote in aptos),
            espaco_estimado_bytes=sum(pacote.espaco_recuperavel_bytes for pacote in aptos),
            data_corte=relatorio.data_corte,
            carregamento_ids=tuple(sorted(pacote.carregamento_id for pacote in aptos)),
        )

    @staticmethod
    def _validar_pacote_antes_exclusao(
        pacote: PacoteRetencaoUnitario,
        arvore: ArvoreCarregamentoRaw,
        data_corte: date,
    ) -> None:
        if pacote.saude == SaudePacote.CRITICO or not pacote.apto_retencao:
            raise RetencaoExecucaoError(f"Pacote {pacote.numero_carregamento} com saude critica.")
        if arvore.data >= data_corte:
            raise RetencaoExecucaoError(f"Carregamento {arvore.numero_carregamento} fora da elegibilidade.")
        if len(arvore.item_ids) != pacote.itens_carregamento:
            raise RetencaoExecucaoError(
                f"Integridade alterada no pacote {pacote.numero_carregamento} (itens)."
            )
        if len(arvore.documento_ids) != pacote.documentos_pdf:
            raise RetencaoExecucaoError(
                f"Integridade alterada no pacote {pacote.numero_carregamento} (PDFs)."
            )

    def _resolver_pdf_dir(self) -> Path:
        if self._pdf_storage_dir is not None:
            return self._pdf_storage_dir
        from infrastructure.database import get_pdf_storage_dir

        return get_pdf_storage_dir()

    def _resolver_xml_dir(self) -> Path:
        if self._xml_storage_dir is not None:
            return self._xml_storage_dir
        from infrastructure.database import get_xml_storage_dir

        return get_xml_storage_dir()

    @staticmethod
    def _remover_arquivos(caminhos: list[str], *, base_dir: Path) -> tuple[int, list[str]]:
        removidos = 0
        falhas: list[str] = []
        vistos: set[str] = set()
        for caminho in caminhos:
            relativo = str(caminho or "").strip()
            if not relativo or relativo in vistos:
                continue
            vistos.add(relativo)
            candidato = Path(relativo)
            absolute = candidato if candidato.is_absolute() else base_dir / relativo
            try:
                if absolute.is_file():
                    absolute.unlink()
                    removidos += 1
            except OSError:
                falhas.append(relativo)
        return removidos, falhas
