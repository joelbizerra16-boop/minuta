from __future__ import annotations

import pandas as pd

from auth.models.usuario import UsuarioPublico
from carregamentos.models.carregamento import (
    Carregamento,
    CarregamentoItem,
    MODALIDADE_VEICULO,
    STATUS_FINALIZADO,
    normalize_chave_nfe,
    normalize_nf_number,
)
from carregamentos.models.fechamento import FechamentoResult
from carregamentos.models.operacional import (
    ClassificacaoOperacionalNf,
    DecisaoOperacional,
    DiagnosticoCarregamento,
    OcorrenciaProcessamentoNf,
    OperacaoProposta,
    PlanoOperacionalLote,
    ResultadoProcessamentoLote,
    SeveridadeDiagnostico,
)
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from carregamentos.services.decisao_operacional_auditoria_service import DecisaoOperacionalAuditoriaService
from carregamentos.services.nf_operacional_classifier import NfOperacionalClassifier, _row_token
from carregamentos.services.validacao_item_carregamento import (
    filtrar_itens_novos_para_insercao,
    montar_lista_itens_pos_complementacao,
    validar_lista_final_itens,
)
from infrastructure.models.constants import (
    AUDIT_CATEGORIA_CARREGAMENTO,
    AUDIT_EVENTO_COMPLEMENTACAO,
    HISTORICO_EVENTO_COMPLEMENTACAO,
)
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.historico_repository import HistoricoRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.repositories.sql.historico_repository import SqlHistoricoRepository
from infrastructure.unit_of_work import UnitOfWork


class ErroEstruturalProcessamento(Exception):
    """Falha que compromete integridade — único caso de interrupção automática."""


class LoteProcessamentoService:
    def __init__(self, repository: SqlCarregamentoRepository) -> None:
        self._repository = repository
        self._classifier = NfOperacionalClassifier()
        self._auditoria_decisao = DecisaoOperacionalAuditoriaService()

    def montar_plano(
        self,
        processed_df: pd.DataFrame,
        diagnostico: DiagnosticoCarregamento,
        decisao_lote: DecisaoOperacional,
        carregamento_existente: Carregamento | None = None,
    ) -> PlanoOperacionalLote:
        if processed_df.empty:
            return PlanoOperacionalLote(
                decisao_lote=decisao_lote,
                severidade=SeveridadeDiagnostico.BLOQUEIO_ESTRUTURAL,
                bloqueio_estrutural=True,
                mensagem_bloqueio="Nenhum dado processado para montar o plano operacional.",
            )

        diagnosticos_nf = self._classifier.classificar_lote(
            processed_df,
            diagnostico,
            decisao_lote,
            carregamento_existente,
        )
        invalidas = [d for d in diagnosticos_nf if d.classificacao == ClassificacaoOperacionalNf.INVALIDA]
        if invalidas and decisao_lote == DecisaoOperacional.NOVO:
            return PlanoOperacionalLote(
                decisao_lote=decisao_lote,
                diagnostico_nf=diagnosticos_nf,
                severidade=SeveridadeDiagnostico.BLOQUEIO_ESTRUTURAL,
                bloqueio_estrutural=True,
                mensagem_bloqueio="; ".join(msg for d in invalidas for msg in d.mensagens),
            )

        inserir = sum(1 for d in diagnosticos_nf if d.acao_proposta == OperacaoProposta.INSERT_ITENS)
        reutilizar = sum(
            1
            for d in diagnosticos_nf
            if d.acao_proposta
            in {OperacaoProposta.REUTILIZAR_REGISTROS, OperacaoProposta.REGISTRAR_REENTREGA}
        )
        return PlanoOperacionalLote(
            decisao_lote=decisao_lote,
            diagnostico_nf=diagnosticos_nf,
            severidade=SeveridadeDiagnostico.INFORMATIVO,
            carregamento_id=diagnostico.carregamento_id,
            itens_para_inserir=inserir,
            nfs_para_inserir=inserir,
            nfs_para_reutilizar=reutilizar,
        )

    def validar_plano_complementacao(
        self,
        carregamento: Carregamento,
        processed_df: pd.DataFrame,
        plano: PlanoOperacionalLote,
    ) -> tuple[list[CarregamentoItem], ResultadoProcessamentoLote]:
        """Validação preventiva antes de qualquer INSERT em complementação."""
        relatorio = ResultadoProcessamentoLote(total_recebidas=len(plano.diagnostico_nf))
        itens_candidatos: list[CarregamentoItem] = []

        for diag_nf in plano.diagnostico_nf:
            if diag_nf.acao_proposta == OperacaoProposta.INSERT_ITENS:
                df_nf = self._filtrar_df_por_token(processed_df, diag_nf.token)
                itens_candidatos.extend(self._build_itens_from_dataframe(df_nf))
                relatorio.processadas += 1
                relatorio.complementadas += 1
                relatorio.ocorrencias.append(
                    OcorrenciaProcessamentoNf(
                        token=diag_nf.token,
                        nf=diag_nf.nf,
                        classificacao=diag_nf.classificacao,
                        sucesso=True,
                        mensagem="Itens planejados para insercao.",
                    )
                )
            elif diag_nf.acao_proposta == OperacaoProposta.REUTILIZAR_REGISTROS:
                relatorio.reentregas += 1
                relatorio.duplicidades += (
                    1 if diag_nf.classificacao == ClassificacaoOperacionalNf.DUPLICIDADE else 0
                )
                relatorio.processadas += 1
                relatorio.ocorrencias.append(
                    OcorrenciaProcessamentoNf(
                        token=diag_nf.token,
                        nf=diag_nf.nf,
                        classificacao=diag_nf.classificacao,
                        sucesso=True,
                        mensagem="Registros reutilizados — sem INSERT.",
                    )
                )
            elif diag_nf.acao_proposta == OperacaoProposta.REGISTRAR_OCORRENCIA:
                relatorio.invalidas += 1
                relatorio.ocorrencias.append(
                    OcorrenciaProcessamentoNf(
                        token=diag_nf.token,
                        nf=diag_nf.nf,
                        classificacao=diag_nf.classificacao,
                        sucesso=False,
                        mensagem="; ".join(diag_nf.mensagens) or "NF invalida.",
                    )
                )
            else:
                relatorio.processadas += 1

        itens_novos, duplicados = filtrar_itens_novos_para_insercao(carregamento, itens_candidatos)
        if duplicados:
            relatorio.duplicidades += len({normalize_nf_number(i.nf) for i in duplicados})

        lista_final, avisos = montar_lista_itens_pos_complementacao(carregamento, itens_novos)
        valido, erros = validar_lista_final_itens(carregamento, lista_final)
        erros_estruturais = [e for e in erros if "Duplicidade detectada" in e]
        if erros_estruturais:
            raise ErroEstruturalProcessamento("; ".join(erros_estruturais))

        if not valido and itens_novos:
            raise ErroEstruturalProcessamento("; ".join(erros) or "Validacao preventiva falhou.")

        for aviso in avisos:
            if "reutilizacao" in aviso:
                continue
            relatorio.ocorrencias.append(
                OcorrenciaProcessamentoNf(
                    token="",
                    nf="",
                    classificacao=ClassificacaoOperacionalNf.DUPLICIDADE,
                    sucesso=True,
                    mensagem=aviso,
                )
            )

        if not itens_novos and relatorio.complementadas == 0 and relatorio.reentregas > 0:
            # Somente reentregas/duplicidades — complementação válida sem INSERT
            return [], relatorio

        return itens_novos, relatorio

    def executar_complementacao(
        self,
        existente: Carregamento,
        *,
        processed_df: pd.DataFrame,
        plano: PlanoOperacionalLote,
        current_user: UsuarioPublico | None,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        ip_origem: str | None,
        motivo_decisao: str = "Operador confirmou complementacao do carregamento.",
    ) -> tuple[Carregamento, ResultadoProcessamentoLote]:
        situacao_anterior = DecisaoOperacionalAuditoriaService.snapshot_carregamento(existente)
        itens_novos, relatorio = self.validar_plano_complementacao(existente, processed_df, plano)

        if not itens_novos:
            if relatorio.reentregas or relatorio.duplicidades:
                # Complementação sem novos itens — apenas impressão/auditoria
                usuario_id = self._resolve_usuario_id(current_user)
                with UnitOfWork() as uow:
                    repo = SqlCarregamentoRepository(uow.session)
                    audit_repo = SqlEventoAuditoriaRepository(uow.session)
                    saved = repo.registrar_impressao(uow.session, existente.id, usuario_id)
                    self._auditoria_decisao.registrar(
                        audit_repo,
                        usuario_id=usuario_id,
                        usuario_nome=DecisaoOperacionalAuditoriaService.usuario_nome(current_user),
                        carregamento_id=saved.id,
                        motivo=motivo_decisao,
                        decisao=DecisaoOperacional.COMPLEMENTAR.value,
                        situacao_anterior=situacao_anterior,
                        situacao_posterior=DecisaoOperacionalAuditoriaService.snapshot_carregamento(saved),
                        nfs_envolvidas=[o.nf for o in relatorio.ocorrencias if o.nf],
                        impactos=[i for d in plano.diagnostico_nf for i in d.impactos],
                        riscos=[r for d in plano.diagnostico_nf for r in d.riscos],
                        recomendacao="Complementacao sem novos itens — reutilizacao de registros.",
                        ip_origem=ip_origem,
                        extras={"relatorio": relatorio.resumo_texto()},
                    )
                reloaded = self._repository.get_by_id(existente.id)
                if reloaded is None:
                    raise ErroEstruturalProcessamento("Carregamento nao encontrado apos complementacao.")
                return reloaded, relatorio
            raise ErroEstruturalProcessamento(
                "Nenhuma NF nova para complementar o carregamento apos validacao preventiva."
            )

        lista_final, _ = montar_lista_itens_pos_complementacao(existente, itens_novos)
        valido, erros = validar_lista_final_itens(existente, lista_final)
        if not valido:
            raise ErroEstruturalProcessamento("; ".join(erros))

        existente.itens = lista_final
        existente.quantidade_nf = len(
            {
                normalize_nf_number(item.nf) or normalize_chave_nfe(item.chave_nfe)
                for item in lista_final
                if normalize_nf_number(item.nf) or normalize_chave_nfe(item.chave_nfe)
            }
        )
        existente.quantidade_itens = len(lista_final)
        existente.peso_total = float(sum(float(item.peso or 0) for item in lista_final))

        usuario_id = self._resolve_usuario_id(current_user)
        with UnitOfWork() as uow:
            repo = SqlCarregamentoRepository(uow.session)
            historico_repo = SqlHistoricoRepository(uow.session)
            audit_repo = SqlEventoAuditoriaRepository(uow.session)

            saved = repo._save_in_session(uow.session, existente)
            if gerar_minuta and not saved.minuta_pdf_path:
                saved.minuta_pdf_path = f"carregamentos/{saved.id}/minuta_carregamento.pdf"
            if gerar_romaneio and not saved.romaneio_pdf_path:
                saved.romaneio_pdf_path = f"carregamentos/{saved.id}/romaneio_entrega.pdf"
            if saved.minuta_pdf_path or saved.romaneio_pdf_path:
                saved = repo._save_in_session(uow.session, saved)

            saved = repo.registrar_impressao(uow.session, saved.id, usuario_id)
            historico_repo.append(
                HistoricoRecord(
                    id=0,
                    carregamento_id=saved.id,
                    usuario_id=usuario_id,
                    evento=HISTORICO_EVENTO_COMPLEMENTACAO,
                    descricao=(
                        f"Complementacao do carregamento {saved.numero_carregamento} "
                        f"com {len(itens_novos)} item(ns) novo(s). {relatorio.resumo_texto()}"
                    ),
                )
            )
            audit_repo.append(
                EventoAuditoriaRecord(
                    id=0,
                    categoria=AUDIT_CATEGORIA_CARREGAMENTO,
                    evento=AUDIT_EVENTO_COMPLEMENTACAO,
                    usuario_id=usuario_id,
                    entidade_tipo="carregamento",
                    entidade_id=saved.id,
                    descricao=f"Complementacao do carregamento {saved.numero_carregamento}",
                    metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                        itens_adicionados=len(itens_novos),
                        gerar_minuta=gerar_minuta,
                        gerar_romaneio=gerar_romaneio,
                        relatorio_lote=relatorio.resumo_texto(),
                    ),
                    ip_origem=ip_origem,
                )
            )
            self._auditoria_decisao.registrar(
                audit_repo,
                usuario_id=usuario_id,
                usuario_nome=DecisaoOperacionalAuditoriaService.usuario_nome(current_user),
                carregamento_id=saved.id,
                motivo=motivo_decisao,
                decisao=DecisaoOperacional.COMPLEMENTAR.value,
                situacao_anterior=situacao_anterior,
                situacao_posterior=DecisaoOperacionalAuditoriaService.snapshot_carregamento(saved),
                nfs_envolvidas=[o.nf for o in relatorio.ocorrencias if o.nf and o.sucesso],
                impactos=[i for d in plano.diagnostico_nf for i in d.impactos],
                riscos=[r for d in plano.diagnostico_nf for r in d.riscos],
                recomendacao="Complementacao validada preventivamente.",
                ip_origem=ip_origem,
                extras={"relatorio": relatorio.resumo_texto()},
            )

        reloaded = self._repository.get_by_id(existente.id)
        if reloaded is None:
            raise ErroEstruturalProcessamento("Carregamento nao encontrado apos complementacao.")
        return reloaded, relatorio

    def validar_itens_novo_carregamento(self, itens: list[CarregamentoItem]) -> None:
        """Valida lista de itens para novo carregamento (sem duplicidade lógica no lote)."""
        carregamento_vazio = Carregamento(
            id=0,
            numero_carregamento="--",
            data="",
            hora="",
            usuario="",
            usuario_id=None,
            motorista="--",
            placa="--",
            filial="",
            data_saida="--",
            quantidade_nf=0,
            quantidade_itens=0,
            peso_total=0,
            status=STATUS_FINALIZADO,
            modalidade=MODALIDADE_VEICULO,
            reentrega=False,
            minuta_pdf_path=None,
            romaneio_pdf_path=None,
            itens=[],
            criado_em="",
        )
        valido, erros = validar_lista_final_itens(carregamento_vazio, itens)
        if not valido:
            raise ErroEstruturalProcessamento("; ".join(erros))

    @staticmethod
    def _filtrar_df_por_token(processed_df: pd.DataFrame, token: str) -> pd.DataFrame:
        if processed_df.empty:
            return processed_df.iloc[0:0]
        mask = processed_df.apply(_row_token, axis=1) == token
        return processed_df[mask].copy()

    @staticmethod
    def _build_itens_from_dataframe(processed_df: pd.DataFrame) -> list[CarregamentoItem]:
        itens: list[CarregamentoItem] = []
        for _, row in processed_df.iterrows():
            itens.append(
                CarregamentoItem(
                    nf=str(row.get("NF", "") or ""),
                    cprod=str(row.get("cProd", "") or ""),
                    descricao=str(row.get("Descricao", "") or ""),
                    quantidade=float(row.get("Qtd", 0) or 0),
                    unidade=str(row.get("Unidade", "") or ""),
                    peso=float(row.get("Peso", 0) or 0),
                    destinatario=str(row.get("Destinatario", "") or ""),
                    rota=str(row.get("ROTA", "") or ""),
                    chave_nfe=str(row.get("ChaveNFe", "") or ""),
                    status_nf=str(row.get("Status", "") or ""),
                )
            )
        return itens

    def _resolve_usuario_id(self, current_user: UsuarioPublico | None) -> int:
        if current_user is not None and current_user.id:
            return int(current_user.id)
        with UnitOfWork() as uow:
            repo = SqlCarregamentoRepository(uow.session)
            return repo._resolve_usuario_id(
                uow.session,
                Carregamento(
                    id=0,
                    numero_carregamento="--",
                    data="",
                    hora="",
                    usuario="sistema",
                    usuario_id=None,
                    motorista="--",
                    placa="--",
                    filial="",
                    data_saida="--",
                    quantidade_nf=0,
                    quantidade_itens=0,
                    peso_total=0,
                    status=STATUS_FINALIZADO,
                    modalidade=MODALIDADE_VEICULO,
                    reentrega=False,
                    minuta_pdf_path=None,
                    romaneio_pdf_path=None,
                    itens=[],
                    criado_em="",
                ),
            )

    @staticmethod
    def resultado_para_mensagem(relatorio: ResultadoProcessamentoLote) -> str:
        return f"Complementacao concluida: {relatorio.resumo_texto()}."

    @staticmethod
    def fechamento_com_relatorio(
        status: str,
        carregamento: Carregamento,
        relatorio: ResultadoProcessamentoLote,
    ) -> FechamentoResult:
        return FechamentoResult(
            status=status,  # type: ignore[arg-type]
            carregamento=carregamento,
            message=LoteProcessamentoService.resultado_para_mensagem(relatorio),
            relatorio_lote=relatorio,
        )
