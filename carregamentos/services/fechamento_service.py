from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from auth.models.usuario import UsuarioPublico
from carregamentos.models.carregamento import (
    MODALIDADE_BALCAO,
    MODALIDADE_VEICULO,
    STATUS_FINALIZADO,
    STATUS_FINALIZADO_R,
    Carregamento,
    CarregamentoItem,
    normalize_chave_nfe,
    normalize_nf_number,
    utc_now_iso,
)
from carregamentos.models.operacional import (
    CenarioOperacional,
    ClassificacaoNfLote,
    DecisaoOperacional,
    DiagnosticoCarregamento,
)
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from carregamentos.services.nf_validation import NfHistoricoValidator, localizar_nf_no_lote
from carregamentos.models.fechamento import FechamentoResult, ImpressaoInfo
from infrastructure.models.constants import (
    AUDIT_CATEGORIA_CARREGAMENTO,
    AUDIT_EVENTO_COMPLEMENTACAO,
    AUDIT_EVENTO_ENTREGA_BALCAO,
    AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
    AUDIT_EVENTO_REENTREGA,
    AUDIT_EVENTO_REIMPRESSAO,
    DOC_TIPO_MINUTA,
    DOC_TIPO_ROMANEIO,
    HISTORICO_EVENTO_COMPLEMENTACAO,
    HISTORICO_EVENTO_ENTREGA_BALCAO,
    HISTORICO_EVENTO_FINALIZACAO,
    HISTORICO_EVENTO_REENTREGA,
)
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord
from infrastructure.repositories.historico_repository import HistoricoRecord
from infrastructure.repositories.sql.evento_auditoria_repository import SqlEventoAuditoriaRepository
from infrastructure.repositories.sql.historico_repository import SqlHistoricoRepository
from infrastructure.unit_of_work import UnitOfWork


class FechamentoCarregamentoService:
    def __init__(self, repository: SqlCarregamentoRepository, data_dir: Path) -> None:
        self._repository = repository
        self._data_dir = data_dir
        self._nf_validator = NfHistoricoValidator(repository)

    @staticmethod
    def _invalidate_analise_cache() -> None:
        from carregamentos.bootstrap import invalidate_analise_operacional_cache

        invalidate_analise_operacional_cache()

    def executar_fechamento_veiculo(
        self,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        *,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        diagnostico: DiagnosticoCarregamento | None = None,
        decisao: DecisaoOperacional | None = None,
        is_reentrega: bool = False,
        confirmar_reimpressao: bool = False,
        ip_origem: str | None = None,
    ) -> FechamentoResult:
        if processed_df.empty:
            return FechamentoResult(status="invalid", message="Nenhum item processado para fechamento.")
        if gerar_minuta is False and gerar_romaneio is False:
            return FechamentoResult(status="invalid", message="Selecione ao menos um documento para impressao.")

        numero = str(summary.get("numero_carga", "") or "").strip()
        if not numero or numero == "--":
            return FechamentoResult(status="invalid", message="Numero de carregamento invalido.")

        if diagnostico is None:
            return FechamentoResult(
                status="invalid",
                message="Analise operacional obrigatoria antes do fechamento.",
            )

        if diagnostico.bloqueia_fechamento:
            mensagem = "; ".join(diagnostico.mensagens) or "Operacao bloqueada pela analise operacional."
            return FechamentoResult(status="invalid", message=mensagem)

        decisao_efetiva = decisao
        if is_reentrega and decisao_efetiva is None:
            decisao_efetiva = DecisaoOperacional.REENTREGA
        if confirmar_reimpressao and decisao_efetiva is None:
            decisao_efetiva = DecisaoOperacional.REIMPRIMIR
        if decisao_efetiva is None:
            if diagnostico.cenario == CenarioOperacional.NOVO:
                decisao_efetiva = DecisaoOperacional.NOVO
            else:
                return FechamentoResult(
                    status="invalid",
                    message="Decisao operacional pendente. Selecione como deseja continuar.",
                )

        if decisao_efetiva == DecisaoOperacional.CANCELAR:
            return FechamentoResult(status="invalid", message="Operacao cancelada pelo operador.")

        return self._executar_decisao_veiculo(
            decisao_efetiva,
            diagnostico=diagnostico,
            summary=summary,
            processed_df=processed_df,
            current_user=current_user,
            gerar_minuta=gerar_minuta,
            gerar_romaneio=gerar_romaneio,
            ip_origem=ip_origem,
        )

    def _executar_decisao_veiculo(
        self,
        decisao: DecisaoOperacional,
        *,
        diagnostico: DiagnosticoCarregamento,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        ip_origem: str | None,
    ) -> FechamentoResult:
        if decisao == DecisaoOperacional.NOVO:
            numero_operacional = self._repository.proximo_numero_carregamento()
            return self._persistir_novo_carregamento(
                numero_carregamento=numero_operacional,
                summary=summary,
                processed_df=processed_df,
                current_user=current_user,
                modalidade=MODALIDADE_VEICULO,
                status=STATUS_FINALIZADO,
                reentrega=False,
                motorista=str(summary.get("motorista", "--") or "--"),
                placa=str(summary.get("placa", "--") or "--"),
                gerar_minuta=gerar_minuta,
                gerar_romaneio=gerar_romaneio,
                ip_origem=ip_origem,
                historico_evento=HISTORICO_EVENTO_FINALIZACAO,
                audit_evento=AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
            )

        carregamento_id = diagnostico.carregamento_id
        if not carregamento_id:
            return FechamentoResult(status="invalid", message="Carregamento de referencia nao identificado.")

        existente = self._repository.get_by_id(int(carregamento_id))
        if existente is None:
            return FechamentoResult(status="invalid", message="Carregamento localizado nao encontrado no banco.")

        if decisao == DecisaoOperacional.REIMPRIMIR:
            return self._tratar_reimpressao(
                existente,
                confirmar_reimpressao=True,
                current_user=current_user,
                ip_origem=ip_origem,
                gerar_minuta=gerar_minuta,
                gerar_romaneio=gerar_romaneio,
            )

        if decisao == DecisaoOperacional.REENTREGA:
            return self._registrar_reentrega_existente(
                existente,
                current_user=current_user,
                ip_origem=ip_origem,
                gerar_minuta=gerar_minuta,
                gerar_romaneio=gerar_romaneio,
            )

        if decisao == DecisaoOperacional.COMPLEMENTAR:
            novos_df = self._filtrar_nfs_novas(processed_df, diagnostico)
            return self._complementar_carregamento(
                existente,
                novos_df=novos_df,
                summary=summary,
                current_user=current_user,
                ip_origem=ip_origem,
                gerar_minuta=gerar_minuta,
                gerar_romaneio=gerar_romaneio,
            )

        return FechamentoResult(status="invalid", message="Decisao operacional invalida para fechamento.")

    @staticmethod
    def _filtrar_nfs_novas(processed_df: pd.DataFrame, diagnostico: DiagnosticoCarregamento) -> pd.DataFrame:
        if processed_df.empty:
            return processed_df.iloc[0:0]
        tokens_novos = {
            item.token
            for item in diagnostico.nfs
            if item.classificacao == ClassificacaoNfLote.NOVA
        }
        if not tokens_novos:
            return processed_df.iloc[0:0]

        def row_token(row: pd.Series) -> str:
            chave = normalize_chave_nfe(row.get("ChaveNFe", ""))
            if chave:
                return chave
            nf_norm = normalize_nf_number(row.get("NF", ""))
            return f"nf:{nf_norm}" if nf_norm else ""

        mask = processed_df.apply(row_token, axis=1).isin(tokens_novos)
        return processed_df[mask].copy()

    def _complementar_carregamento(
        self,
        existente: Carregamento,
        *,
        novos_df: pd.DataFrame,
        summary: dict[str, Any],
        current_user: UsuarioPublico | None,
        ip_origem: str | None,
        gerar_minuta: bool,
        gerar_romaneio: bool,
    ) -> FechamentoResult:
        novos_itens = self._build_itens_from_dataframe(novos_df)
        if not novos_itens:
            return FechamentoResult(status="invalid", message="Nenhuma NF nova para complementar o carregamento.")

        chaves_existentes = {
            (
                normalize_nf_number(item.nf),
                normalize_chave_nfe(item.chave_nfe),
                str(item.cprod or "").strip(),
            )
            for item in existente.itens
        }
        itens_para_inserir: list[CarregamentoItem] = []
        for item in novos_itens:
            chave_item = (
                normalize_nf_number(item.nf),
                normalize_chave_nfe(item.chave_nfe),
                str(item.cprod or "").strip(),
            )
            if chave_item in chaves_existentes:
                continue
            itens_para_inserir.append(item)

        if not itens_para_inserir:
            return FechamentoResult(
                status="invalid",
                message="Todas as NFs novas ja existem no carregamento. Nenhuma complementacao realizada.",
            )

        todos_itens = list(existente.itens) + itens_para_inserir
        existente.itens = todos_itens
        existente.quantidade_nf = len(
            {
                normalize_nf_number(item.nf) or normalize_chave_nfe(item.chave_nfe)
                for item in todos_itens
                if normalize_nf_number(item.nf) or normalize_chave_nfe(item.chave_nfe)
            }
        )
        existente.quantidade_itens = len(todos_itens)
        existente.peso_total = float(sum(float(item.peso or 0) for item in todos_itens))

        usuario_id = self._resolve_usuario_id(current_user)
        try:
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
                            f"com {len(itens_para_inserir)} item(ns) novo(s)."
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
                            itens_adicionados=len(itens_para_inserir),
                            gerar_minuta=gerar_minuta,
                            gerar_romaneio=gerar_romaneio,
                        ),
                        ip_origem=ip_origem,
                    )
                )
        except Exception as exc:
            return FechamentoResult(status="error", message=f"Falha ao complementar carregamento: {exc}")

        reloaded = self._repository.get_by_id(existente.id)
        if reloaded is None:
            return FechamentoResult(status="error", message="Carregamento nao encontrado apos complementacao.")
        self._invalidate_analise_cache()
        return FechamentoResult(status="complementacao", carregamento=reloaded)

    def _registrar_reentrega_existente(
        self,
        existente: Carregamento,
        *,
        current_user: UsuarioPublico | None,
        ip_origem: str | None,
        gerar_minuta: bool,
        gerar_romaneio: bool,
    ) -> FechamentoResult:
        existente.status = STATUS_FINALIZADO_R
        existente.reentrega = True
        usuario_id = self._resolve_usuario_id(current_user)
        try:
            with UnitOfWork() as uow:
                repo = SqlCarregamentoRepository(uow.session)
                historico_repo = SqlHistoricoRepository(uow.session)
                audit_repo = SqlEventoAuditoriaRepository(uow.session)

                saved = repo._save_in_session(uow.session, existente)
                saved = repo.registrar_impressao(uow.session, saved.id, usuario_id)
                historico_repo.append(
                    HistoricoRecord(
                        id=0,
                        carregamento_id=saved.id,
                        usuario_id=usuario_id,
                        evento=HISTORICO_EVENTO_REENTREGA,
                        descricao=f"Reentrega registrada para o carregamento {saved.numero_carregamento}.",
                    )
                )
                audit_repo.append(
                    EventoAuditoriaRecord(
                        id=0,
                        categoria=AUDIT_CATEGORIA_CARREGAMENTO,
                        evento=AUDIT_EVENTO_REENTREGA,
                        usuario_id=usuario_id,
                        entidade_tipo="carregamento",
                        entidade_id=saved.id,
                        descricao=f"Reentrega do carregamento {saved.numero_carregamento}",
                        metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                            gerar_minuta=gerar_minuta,
                            gerar_romaneio=gerar_romaneio,
                        ),
                        ip_origem=ip_origem,
                    )
                )
        except Exception as exc:
            return FechamentoResult(status="error", message=f"Falha ao registrar reentrega: {exc}")

        reloaded = self._repository.get_by_id(existente.id)
        if reloaded is None:
            return FechamentoResult(status="error", message="Carregamento nao encontrado apos reentrega.")
        self._invalidate_analise_cache()
        return FechamentoResult(status="reimpressao", carregamento=reloaded)

    def executar_fechamento_balcao(
        self,
        termo_busca: str,
        summary: dict[str, Any],
        lookup_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        *,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        is_reentrega: bool = False,
        confirmar_reimpressao: bool = False,
        standalone_balcao: bool = True,
        ip_origem: str | None = None,
    ) -> FechamentoResult:
        nf_df = localizar_nf_no_lote(lookup_df, termo_busca)
        if nf_df.empty:
            return FechamentoResult(status="invalid", message="NF nao encontrada nos XMLs carregados no sistema.")

        numero_operacional = self._repository.proximo_numero_carregamento()

        if standalone_balcao:
            balcao_summary = {
                "filial": str(summary.get("filial", "BRIDA") or "BRIDA"),
                "data_saida": str(summary.get("data_saida", "--") or "--"),
                "nf_count": int(nf_df["NF"].nunique()),
                "item_count": int(len(nf_df)),
                "peso_total": float(nf_df["Peso"].sum()),
            }
        else:
            balcao_summary = dict(summary)
            balcao_summary["nf_count"] = int(nf_df["NF"].nunique())
            balcao_summary["item_count"] = int(len(nf_df))
            balcao_summary["peso_total"] = float(nf_df["Peso"].sum())

        return self._persistir_novo_carregamento(
            numero_carregamento=numero_operacional,
            summary=balcao_summary,
            processed_df=nf_df,
            current_user=current_user,
            modalidade=MODALIDADE_BALCAO,
            status=STATUS_FINALIZADO_R if is_reentrega else STATUS_FINALIZADO,
            reentrega=is_reentrega,
            motorista="--",
            placa="--",
            gerar_minuta=gerar_minuta,
            gerar_romaneio=gerar_romaneio,
            ip_origem=ip_origem,
            historico_evento=HISTORICO_EVENTO_REENTREGA if is_reentrega else HISTORICO_EVENTO_ENTREGA_BALCAO,
            audit_evento=AUDIT_EVENTO_REENTREGA if is_reentrega else AUDIT_EVENTO_ENTREGA_BALCAO,
        )

    def gravar_pdfs_pos_commit(
        self,
        carregamento: Carregamento,
        *,
        minuta_pdf: bytes | None,
        romaneio_pdf: bytes | None,
    ) -> Carregamento:
        storage_folder = self._repository.storage_dir / str(carregamento.id)
        storage_folder.mkdir(parents=True, exist_ok=True)
        hashes: dict[str, str] = {}

        if minuta_pdf:
            target = storage_folder / "minuta_carregamento.pdf"
            target.write_bytes(minuta_pdf)
            carregamento.minuta_pdf_path = f"carregamentos/{carregamento.id}/minuta_carregamento.pdf"
            hashes[DOC_TIPO_MINUTA] = hashlib.sha256(minuta_pdf).hexdigest()

        if romaneio_pdf:
            target = storage_folder / "romaneio_entrega.pdf"
            target.write_bytes(romaneio_pdf)
            carregamento.romaneio_pdf_path = f"carregamentos/{carregamento.id}/romaneio_entrega.pdf"
            hashes[DOC_TIPO_ROMANEIO] = hashlib.sha256(romaneio_pdf).hexdigest()

        if hashes:
            with UnitOfWork() as uow:
                repo = SqlCarregamentoRepository(uow.session)
                repo.sync_document_hashes(uow.session, carregamento.id, hashes)

        return carregamento

    def executar_reimpressao(
        self,
        carregamento_id: int,
        current_user: UsuarioPublico | None,
        *,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        ip_origem: str | None = None,
    ) -> FechamentoResult:
        existente = self._repository.get_by_id(carregamento_id)
        if existente is None:
            return FechamentoResult(status="invalid", message="Carregamento nao encontrado.")
        return self._tratar_reimpressao(
            existente,
            confirmar_reimpressao=True,
            current_user=current_user,
            ip_origem=ip_origem,
            gerar_minuta=gerar_minuta,
            gerar_romaneio=gerar_romaneio,
        )

    def _tratar_reimpressao(
        self,
        existente: Carregamento,
        *,
        confirmar_reimpressao: bool,
        current_user: UsuarioPublico | None,
        ip_origem: str | None,
        gerar_minuta: bool,
        gerar_romaneio: bool,
    ) -> FechamentoResult:
        info = ImpressaoInfo(
            carregamento_id=existente.id,
            numero_carregamento=existente.numero_carregamento,
            primeira_impressao_data=self._formatar_data_hora(existente.data, existente.hora),
            primeira_impressao_usuario=existente.usuario,
            quantidade_impressoes=max(int(existente.quantidade_impressoes or 0), 1),
        )
        if not confirmar_reimpressao:
            return FechamentoResult(
                status="needs_reimpressao_confirm",
                carregamento=existente,
                impressao_info=info,
            )

        usuario_id = self._resolve_usuario_id(current_user)
        with UnitOfWork() as uow:
            repo = SqlCarregamentoRepository(uow.session)
            audit_repo = SqlEventoAuditoriaRepository(uow.session)
            atualizado = repo.registrar_impressao(uow.session, existente.id, usuario_id)
            audit_repo.append(
                EventoAuditoriaRecord(
                    id=0,
                    categoria=AUDIT_CATEGORIA_CARREGAMENTO,
                    evento=AUDIT_EVENTO_REIMPRESSAO,
                    usuario_id=usuario_id,
                    entidade_tipo="carregamento",
                    entidade_id=atualizado.id,
                    descricao=f"Reimpressao do carregamento {atualizado.numero_carregamento}",
                    metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                        gerar_minuta=gerar_minuta,
                        gerar_romaneio=gerar_romaneio,
                        quantidade_impressoes=atualizado.quantidade_impressoes,
                    ),
                    ip_origem=ip_origem,
                )
            )

        self._invalidate_analise_cache()
        return FechamentoResult(status="reimpressao", carregamento=atualizado, impressao_info=info)

    def _persistir_novo_carregamento(
        self,
        *,
        numero_carregamento: str,
        summary: dict[str, Any],
        processed_df: pd.DataFrame,
        current_user: UsuarioPublico | None,
        modalidade: str,
        status: str,
        reentrega: bool,
        motorista: str,
        placa: str,
        gerar_minuta: bool,
        gerar_romaneio: bool,
        ip_origem: str | None,
        historico_evento: str,
        audit_evento: str,
    ) -> FechamentoResult:
        now = datetime.now()
        itens = self._build_itens_from_dataframe(processed_df)
        usuario_login = str(current_user.usuario if current_user else "sistema")
        usuario_id = self._resolve_usuario_id(current_user)

        carregamento = Carregamento(
            id=0,
            numero_carregamento=numero_carregamento,
            data=now.strftime("%Y-%m-%d"),
            hora=now.strftime("%H:%M:%S"),
            usuario=usuario_login,
            usuario_id=usuario_id,
            motorista=motorista,
            placa=placa,
            filial=str(summary.get("filial", "BRIDA") or "BRIDA"),
            data_saida=str(summary.get("data_saida", "--") or "--"),
            quantidade_nf=int(summary.get("nf_count", 0) or 0),
            quantidade_itens=int(summary.get("item_count", 0) or len(itens)),
            peso_total=float(summary.get("peso_total", 0) or 0),
            status=status,
            modalidade=modalidade,
            reentrega=reentrega,
            minuta_pdf_path=None,
            romaneio_pdf_path=None,
            itens=itens,
            criado_em=utc_now_iso(),
        )

        try:
            saved: Carregamento | None = None
            with UnitOfWork() as uow:
                repo = SqlCarregamentoRepository(uow.session)
                historico_repo = SqlHistoricoRepository(uow.session)
                audit_repo = SqlEventoAuditoriaRepository(uow.session)

                saved = repo._save_in_session(uow.session, carregamento)
                if gerar_minuta:
                    saved.minuta_pdf_path = f"carregamentos/{saved.id}/minuta_carregamento.pdf"
                else:
                    saved.minuta_pdf_path = None
                if gerar_romaneio:
                    saved.romaneio_pdf_path = f"carregamentos/{saved.id}/romaneio_entrega.pdf"
                else:
                    saved.romaneio_pdf_path = None
                if saved.minuta_pdf_path or saved.romaneio_pdf_path:
                    saved = repo._save_in_session(uow.session, saved)

                saved = repo.registrar_impressao(uow.session, saved.id, usuario_id)

                historico_repo.append(
                    HistoricoRecord(
                        id=0,
                        carregamento_id=saved.id,
                        usuario_id=usuario_id,
                        evento=historico_evento,
                        descricao=f"Fechamento do carregamento {saved.numero_carregamento}",
                    )
                )
                audit_repo.append(
                    EventoAuditoriaRecord(
                        id=0,
                        categoria=AUDIT_CATEGORIA_CARREGAMENTO,
                        evento=audit_evento,
                        usuario_id=usuario_id,
                        entidade_tipo="carregamento",
                        entidade_id=saved.id,
                        descricao=f"Primeira impressao do carregamento {saved.numero_carregamento}",
                        metadados_json=SqlEventoAuditoriaRepository.build_metadados(
                            gerar_minuta=gerar_minuta,
                            gerar_romaneio=gerar_romaneio,
                            reentrega=reentrega,
                            modalidade=modalidade,
                        ),
                        ip_origem=ip_origem,
                    )
                )
        except Exception as exc:
            return FechamentoResult(status="error", message=f"Falha ao salvar carregamento: {exc}")

        reloaded = self._repository.get_by_id(saved.id)
        if reloaded is None:
            return FechamentoResult(status="error", message="Carregamento nao encontrado apos commit.")
        self._invalidate_analise_cache()
        return FechamentoResult(status="primeira_impressao", carregamento=reloaded)

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
                    criado_em=utc_now_iso(),
                ),
            )

    @staticmethod
    def _formatar_data_hora(data: str, hora: str) -> str:
        parts = str(data or "").split("-")
        if len(parts) == 3:
            return f"{parts[2]}/{parts[1]}/{parts[0]} {hora}"
        return f"{data} {hora}".strip()
