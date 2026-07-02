from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal

import pandas as pd
import streamlit as st

from auth.security.session import get_current_user
from carregamentos.bootstrap import (
    get_analise_operacional_service,
    get_fechamento_service,
    invalidate_analise_operacional_cache,
)
from carregamentos.models.carregamento import NfHistoricoConflito
from carregamentos.models.fechamento import FechamentoResult
from carregamentos.models.operacional import DecisaoOperacional, DiagnosticoCarregamento
from carregamentos.services.nf_validation import localizar_nf_no_lote
from utils.streamlit_tables import build_table_column_config

FinalizeStatus = Literal[
    "saved",
    "already",
    "invalid",
    "needs_reentrega",
    "cancelled",
    "balcao_saved",
    "balcao_not_found",
    "balcao_needs_reentrega",
    "balcao_preview",
    "needs_reimpressao_confirm",
    "reimpressao",
    "complementacao",
    "error",
]

DECISAO_OPERACIONAL_LABELS: dict[DecisaoOperacional, str] = {
    DecisaoOperacional.NOVO: "Registrar novo carregamento",
    DecisaoOperacional.REIMPRIMIR: "Reimprimir documentos",
    DecisaoOperacional.COMPLEMENTAR: "Complementar carregamento",
    DecisaoOperacional.REENTREGA: "Registrar reentrega",
    DecisaoOperacional.CANCELAR: "Cancelar operacao",
}

OPERACIONAL_CONTEXT_KEYS = (
    "operacional_diagnostico",
    "operacional_decisao",
    "operacional_excel_nome",
    "carregamento_saved_id",
    "carregamento_fechado",
    "carregamento_finalize_error",
    "carregamento_finalize_message",
    "carregamento_finalize_warning",
    "carregamento_finalized_signature",
    "pdf_download_payload",
    "pdf_download_name",
    "pdf_download_mime",
    "_prepared_processed_df",
    "_display_table_df",
    "_processed_data_version",
)


@dataclass(frozen=True)
class PdfDownloadPackage:
    payload: bytes
    file_name: str
    mime_type: str


def _store_reentrega_pending(
    conflitos: list[NfHistoricoConflito],
    contexto: str,
    *,
    mensagem_balcao: bool = False,
) -> None:
    st.session_state["reentrega_pending"] = True
    st.session_state["reentrega_contexto"] = contexto
    if mensagem_balcao:
        st.session_state["reentrega_conflitos"] = [
            conflito.formatar_mensagem_balcao() for conflito in conflitos
        ]
    else:
        st.session_state["reentrega_conflitos"] = [conflito.formatar_mensagem() for conflito in conflitos]


def clear_reentrega_pending() -> None:
    st.session_state.pop("reentrega_pending", None)
    st.session_state.pop("reentrega_contexto", None)
    st.session_state.pop("reentrega_conflitos", None)
    st.session_state.pop("reentrega_balcao_termo", None)


def clear_balcao_pending() -> None:
    st.session_state.pop("balcao_pending_confirm", None)
    st.session_state.pop("balcao_confirm_termo", None)
    st.session_state.pop("balcao_force_reentrega", None)


def clear_reimpressao_pending() -> None:
    st.session_state.pop("reimpressao_pending", None)
    st.session_state.pop("reimpressao_info", None)


def cancelar_operacao_pendente() -> None:
    st.session_state.pop("operacional_decisao", None)
    st.session_state.pop("pdf_download_payload", None)
    st.session_state.pop("pdf_download_name", None)
    st.session_state.pop("pdf_download_mime", None)
    st.session_state.pop("carregamento_finalize_message", None)
    st.session_state.pop("carregamento_finalize_warning", None)
    st.session_state.pop("carregamento_finalize_error", None)


def clear_contexto_operacional() -> None:
    clear_reentrega_pending()
    clear_reimpressao_pending()
    clear_balcao_pending()
    for key in OPERACIONAL_CONTEXT_KEYS:
        st.session_state.pop(key, None)
    for key in list(st.session_state.keys()):
        if str(key).startswith("historico_carregamento_loaded_") or str(key).startswith("auditoria_nf_loaded_") or str(key).startswith("auditoria_nf_expanded_") or str(key).startswith("auditoria_nf_lista_") or str(key).startswith("auditoria_nf_foco_decisao_"):
            st.session_state.pop(key, None)
    try:
        from carregamentos.ui.auditoria_nf_panel import _carregar_auditoria_nf_cache, _carregar_extrato_nf_cache
        from carregamentos.ui.historico_carregamento_panel import _carregar_painel_cache

        _carregar_painel_cache.clear()
        _carregar_auditoria_nf_cache.clear()
        _carregar_extrato_nf_cache.clear()
    except Exception:
        pass


def get_operacional_diagnostico() -> DiagnosticoCarregamento | None:
    payload = st.session_state.get("operacional_diagnostico")
    if isinstance(payload, dict):
        return DiagnosticoCarregamento.from_dict(payload)
    return None


def set_operacional_diagnostico(diagnostico: DiagnosticoCarregamento) -> None:
    st.session_state["operacional_diagnostico"] = diagnostico.to_dict()
    st.session_state.pop("operacional_decisao", None)


def get_operacional_decisao() -> DecisaoOperacional | None:
    value = st.session_state.get("operacional_decisao")
    if not value:
        return None
    try:
        return DecisaoOperacional(str(value))
    except ValueError:
        return None


def set_operacional_decisao(decisao: DecisaoOperacional) -> None:
    st.session_state["operacional_decisao"] = decisao.value


def executar_analise_operacional(processed_df: pd.DataFrame) -> DiagnosticoCarregamento:
    diagnostico = get_analise_operacional_service().analisar_lote_processado(processed_df)
    set_operacional_diagnostico(diagnostico)
    return diagnostico


def queue_processing_action(action: dict[str, object]) -> None:
    st.session_state["_processing_action"] = action


def resolve_operational_panel_mode(*, has_excel_loaded: bool, processed_df_empty: bool) -> str:
    if st.session_state.get("reentrega_pending"):
        return "reentrega"
    if st.session_state.get("reimpressao_pending"):
        return "reimpressao"
    if st.session_state.get("balcao_pending_confirm"):
        return "balcao_confirm"

    diagnostico = get_operacional_diagnostico()
    if has_excel_loaded and not processed_df_empty and diagnostico is not None:
        if diagnostico.bloqueia_fechamento and diagnostico.requer_decisao:
            return "carregamento_bloqueado"
        if diagnostico.requer_decisao and get_operacional_decisao() is None:
            return "carregamento_localizado"

    if not has_excel_loaded:
        return "balcao"
    if has_excel_loaded and not processed_df_empty:
        return "fechamento"
    return "idle"


def sync_processing_context_for_excel(has_excel_loaded: bool) -> None:
    if has_excel_loaded:
        clear_balcao_pending()
    else:
        st.session_state.pop("carregamento_saved_id", None)
        st.session_state.pop("carregamento_fechado", None)


def on_processing_panel_primary_click() -> None:
    mode = str(st.session_state.get("_operational_panel_mode", "idle") or "idle")
    if mode == "reentrega":
        queue_processing_action({"type": "reentrega_confirm"})
    elif mode == "reimpressao":
        queue_processing_action({"type": "baixar_pdf", "confirmar_reimpressao": True})
    elif mode == "balcao_confirm":
        queue_processing_action({"type": "balcao_confirm"})
    elif mode == "balcao":
        queue_processing_action(
            {
                "type": "balcao_iniciar",
                "termo": str(st.session_state.get("entrega_balcao_termo", "") or ""),
            }
        )
    elif mode == "carregamento_bloqueado":
        queue_processing_action({"type": "operacional_cancel"})
    elif mode == "fechamento":
        queue_processing_action({"type": "finalizar_carregamento"})


def on_processing_panel_secondary_click() -> None:
    mode = str(st.session_state.get("_operational_panel_mode", "idle") or "idle")
    if mode == "reentrega":
        queue_processing_action({"type": "reentrega_cancel"})
    elif mode == "reimpressao":
        queue_processing_action({"type": "reimpressao_cancel"})
    elif mode == "balcao_confirm":
        queue_processing_action({"type": "balcao_cancel"})
    elif mode in {"carregamento_localizado", "carregamento_bloqueado"}:
        queue_processing_action({"type": "operacional_cancel"})


def on_baixar_pdf_click() -> None:
    queue_processing_action({"type": "baixar_pdf"})


def render_balcao_nf_preview(lookup_df: pd.DataFrame, termo: str) -> None:
    nf_df = localizar_nf_no_lote(lookup_df, termo)
    if nf_df.empty:
        return

    preview_df = nf_df.copy()
    rename_map = {
        "cProd": "Produto",
        "Qtd": "Quantidade",
    }
    preview_df = preview_df.rename(columns=rename_map)
    display_columns = [
        column
        for column in ["NF", "Destinatario", "Produto", "Descricao", "Quantidade", "Peso", "ROTA"]
        if column in preview_df.columns
    ]

    st.dataframe(
        preview_df[display_columns],
        width="stretch",
        hide_index=True,
        column_config=build_table_column_config(preview_df[display_columns]),
        row_height=56,
    )


def iniciar_entrega_balcao(
    termo_busca: str,
    lookup_df: pd.DataFrame,
    *,
    standalone_balcao: bool = True,
) -> FinalizeStatus:
    termo = str(termo_busca or "").strip()
    if lookup_df.empty or not termo:
        return "balcao_not_found"

    nf_df = localizar_nf_no_lote(lookup_df, termo)
    if nf_df.empty:
        return "balcao_not_found"

    from carregamentos.bootstrap import get_carregamento_service

    service = get_carregamento_service()
    conflitos = service.validar_conflitos_nf(nf_df)
    if conflitos:
        _store_reentrega_pending(conflitos, "balcao", mensagem_balcao=standalone_balcao)
        st.session_state["reentrega_balcao_termo"] = termo
        return "balcao_needs_reentrega"

    st.session_state["balcao_pending_confirm"] = True
    st.session_state["balcao_confirm_termo"] = termo
    return "balcao_preview"


def _apply_fechamento_result(result: FechamentoResult) -> FinalizeStatus:
    if result.status == "needs_reentrega":
        _store_reentrega_pending(list(result.conflitos), "veiculo")
        return "needs_reentrega"
    if result.status == "needs_reimpressao_confirm" and result.impressao_info is not None:
        st.session_state["reimpressao_pending"] = True
        st.session_state["reimpressao_info"] = result.impressao_info
        return "needs_reimpressao_confirm"
    if result.status == "invalid":
        st.session_state["carregamento_finalize_error"] = result.message or "Dados invalidos para fechamento."
        return "invalid"
    if result.status == "error":
        st.session_state["carregamento_finalize_error"] = result.message or "Falha ao salvar no banco de dados."
        return "error"
    if result.carregamento is None:
        st.session_state["carregamento_finalize_error"] = "Carregamento nao retornado apos fechamento."
        return "invalid"

    clear_reentrega_pending()
    clear_reimpressao_pending()
    st.session_state["carregamento_saved_id"] = result.carregamento.id
    st.session_state["carregamento_fechado"] = result.carregamento
    invalidate_analise_operacional_cache()
    if result.status == "complementacao":
        return "complementacao"
    if result.status == "reimpressao":
        return "reimpressao"
    return "saved"


def executar_fechamento_veiculo_para_pdf(
    summary: dict[str, Any],
    processed_df: pd.DataFrame,
    *,
    gerar_minuta: bool,
    gerar_romaneio: bool,
    force_reentrega: bool = False,
    confirmar_reimpressao: bool = False,
    diagnostico: DiagnosticoCarregamento | None = None,
    decisao: DecisaoOperacional | None = None,
) -> tuple[FinalizeStatus, FechamentoResult | None]:
    diagnostico_efetivo = diagnostico or get_operacional_diagnostico()
    decisao_efetiva = decisao or get_operacional_decisao()
    if force_reentrega and decisao_efetiva is None:
        decisao_efetiva = DecisaoOperacional.REENTREGA
    if confirmar_reimpressao and decisao_efetiva is None:
        decisao_efetiva = DecisaoOperacional.REIMPRIMIR

    fechamento = get_fechamento_service()
    result = fechamento.executar_fechamento_veiculo(
        summary=summary,
        processed_df=processed_df,
        current_user=get_current_user(),
        gerar_minuta=gerar_minuta,
        gerar_romaneio=gerar_romaneio,
        diagnostico=diagnostico_efetivo,
        decisao=decisao_efetiva,
        is_reentrega=force_reentrega,
        confirmar_reimpressao=confirmar_reimpressao,
    )
    return _apply_fechamento_result(result), result


def executar_fechamento_balcao_para_pdf(
    termo_busca: str,
    summary: dict[str, Any],
    lookup_df: pd.DataFrame,
    *,
    gerar_minuta: bool,
    gerar_romaneio: bool,
    force_reentrega: bool = False,
    confirmar_reimpressao: bool = False,
    standalone_balcao: bool = True,
) -> tuple[FinalizeStatus, FechamentoResult | None]:
    fechamento = get_fechamento_service()
    result = fechamento.executar_fechamento_balcao(
        termo_busca=termo_busca,
        summary=summary,
        lookup_df=lookup_df,
        current_user=get_current_user(),
        gerar_minuta=gerar_minuta,
        gerar_romaneio=gerar_romaneio,
        is_reentrega=force_reentrega,
        confirmar_reimpressao=confirmar_reimpressao,
        standalone_balcao=standalone_balcao,
    )
    status = _apply_fechamento_result(result)
    if status == "saved":
        return "balcao_saved", result
    if status == "needs_reentrega":
        _store_reentrega_pending(list(result.conflitos), "balcao", mensagem_balcao=standalone_balcao)
        st.session_state["reentrega_balcao_termo"] = termo_busca
        return "balcao_needs_reentrega", result
    if status in {"reimpressao", "complementacao"}:
        return "balcao_saved", result
    return status, result


def persistir_pdfs_apos_fechamento(
    carregamento,
    *,
    minuta_pdf: bytes | None,
    romaneio_pdf: bytes | None,
):
    return get_fechamento_service().gravar_pdfs_pos_commit(
        carregamento,
        minuta_pdf=minuta_pdf,
        romaneio_pdf=romaneio_pdf,
    )


# Compatibilidade com fluxos legados que ainda chamam finalizacao direta.
def finalize_carregamento_operacional(
    summary: dict[str, Any],
    processed_df: pd.DataFrame,
    carregamento_pdf: bytes,
    romaneio_pdf: bytes | None,
    *,
    force_reentrega: bool = False,
) -> tuple[int | None, FinalizeStatus]:
    diagnostico = executar_analise_operacional(processed_df)
    decisao = DecisaoOperacional.REENTREGA if force_reentrega else None
    if decisao is None and not diagnostico.requer_decisao:
        decisao = DecisaoOperacional.NOVO
    status, result = executar_fechamento_veiculo_para_pdf(
        summary=summary,
        processed_df=processed_df,
        gerar_minuta=bool(carregamento_pdf),
        gerar_romaneio=bool(romaneio_pdf),
        force_reentrega=force_reentrega,
        confirmar_reimpressao=force_reentrega,
        diagnostico=diagnostico,
        decisao=decisao,
    )
    if status in {"saved", "reimpressao", "complementacao", "balcao_saved"} and result and result.carregamento:
        persistir_pdfs_apos_fechamento(
            result.carregamento,
            minuta_pdf=carregamento_pdf or None,
            romaneio_pdf=romaneio_pdf,
        )
        return result.carregamento.id, status
    if status == "already":
        saved_id = st.session_state.get("carregamento_saved_id")
        return saved_id, status
    return None, status


def registrar_entrega_balcao(
    termo_busca: str,
    summary: dict[str, Any],
    lookup_df: pd.DataFrame,
    carregamento_pdf: bytes | None,
    romaneio_pdf: bytes | None,
    *,
    force_reentrega: bool = False,
    standalone_balcao: bool = True,
) -> tuple[int | None, FinalizeStatus]:
    status, result = executar_fechamento_balcao_para_pdf(
        termo_busca=termo_busca,
        summary=summary,
        lookup_df=lookup_df,
        gerar_minuta=bool(carregamento_pdf),
        gerar_romaneio=bool(romaneio_pdf),
        force_reentrega=force_reentrega,
        confirmar_reimpressao=force_reentrega,
        standalone_balcao=standalone_balcao,
    )
    if status in {"balcao_saved", "reimpressao", "complementacao"} and result and result.carregamento:
        persistir_pdfs_apos_fechamento(
            result.carregamento,
            minuta_pdf=carregamento_pdf or None,
            romaneio_pdf=romaneio_pdf,
        )
        clear_balcao_pending()
        return result.carregamento.id, status
    return None, status
