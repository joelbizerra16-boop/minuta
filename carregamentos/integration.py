from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
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
from carregamentos.models.operacional import (
    CenarioOperacional,
    DecisaoOperacional,
    DiagnosticoCarregamento,
)
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

OPERACIONAL_DECISAO_WIDGET_KEY = "operacional_decisao_widget"
OPERACIONAL_ANALISE_CONFIRMADA_KEY = "operacional_analise_confirmada"
OPERACIONAL_CONTINUACAO_AUDITORIA_KEY = "operacional_continuacao_auditoria"
OPERACIONAL_CONTINUAR_HISTORICO_VALUE = "CONTINUAR_HISTORICO"

_OPERACIONAL_HISTORICO_LOGGER = logging.getLogger("minuta.operacional.historico")

OPERACIONAL_CONTEXT_KEYS = (
    "operacional_diagnostico",
    "operacional_decisao",
    OPERACIONAL_DECISAO_WIDGET_KEY,
    OPERACIONAL_ANALISE_CONFIRMADA_KEY,
    OPERACIONAL_CONTINUACAO_AUDITORIA_KEY,
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


def is_operacional_analise_confirmada() -> bool:
    return bool(st.session_state.get(OPERACIONAL_ANALISE_CONFIRMADA_KEY))


def inferir_decisao_operacional(diagnostico: DiagnosticoCarregamento) -> DecisaoOperacional:
    cenario = diagnostico.cenario
    if cenario == CenarioOperacional.REIMPRESSAO:
        return DecisaoOperacional.REIMPRIMIR
    if cenario == CenarioOperacional.COMPLEMENTACAO:
        return DecisaoOperacional.COMPLEMENTAR
    if cenario == CenarioOperacional.REENTREGA:
        return DecisaoOperacional.REENTREGA
    if cenario == CenarioOperacional.NF_CANCELADA:
        return DecisaoOperacional.CANCELAR
    if DecisaoOperacional.REIMPRIMIR in diagnostico.opcoes_decisao:
        return DecisaoOperacional.REIMPRIMIR
    if DecisaoOperacional.COMPLEMENTAR in diagnostico.opcoes_decisao:
        return DecisaoOperacional.COMPLEMENTAR
    if DecisaoOperacional.REENTREGA in diagnostico.opcoes_decisao:
        return DecisaoOperacional.REENTREGA
    if DecisaoOperacional.NOVO in diagnostico.opcoes_decisao:
        return DecisaoOperacional.NOVO
    return DecisaoOperacional.NOVO


def requer_confirmacao_explicita_historico(diagnostico: DiagnosticoCarregamento) -> bool:
    return (
        diagnostico.bloqueia_fechamento
        and diagnostico.opcoes_decisao == [DecisaoOperacional.CANCELAR]
        and diagnostico.cenario != CenarioOperacional.NF_CANCELADA
    )


def _registrar_continuacao_historico(diagnostico: DiagnosticoCarregamento) -> None:
    agora = datetime.now()
    usuario = get_current_user()
    nome_usuario = str(usuario.nome if usuario and usuario.nome else "--")
    summary = st.session_state.get("summary") or {}
    carregamento_atual = str(summary.get("numero_carga", "") or "--")
    registro = {
        "usuario": nome_usuario,
        "data": agora.strftime("%d/%m/%Y"),
        "hora": agora.strftime("%H:%M:%S"),
        "carregamento_atual": carregamento_atual,
        "quantidade_nfs_historico": int(diagnostico.nfs_existentes),
        "nfs_reutilizadas": int(diagnostico.nfs_existentes),
        "nfs_novas": int(diagnostico.nfs_novas),
        "decisao": (
            "Operador autorizou manualmente o processamento apos analise do historico operacional."
        ),
    }
    st.session_state[OPERACIONAL_CONTINUACAO_AUDITORIA_KEY] = registro
    _OPERACIONAL_HISTORICO_LOGGER.info(
        "%s usuario=%s data=%s hora=%s carregamento=%s nfs_historico=%s nfs_novas=%s",
        registro["decisao"],
        nome_usuario,
        registro["data"],
        registro["hora"],
        carregamento_atual,
        registro["quantidade_nfs_historico"],
        registro["nfs_novas"],
    )


def confirmar_analise_operacional_continuacao() -> None:
    st.session_state[OPERACIONAL_ANALISE_CONFIRMADA_KEY] = True
    st.session_state.pop("operacional_decisao", None)
    st.session_state.pop(OPERACIONAL_DECISAO_WIDGET_KEY, None)


def get_diagnostico_efetivo_fechamento() -> DiagnosticoCarregamento | None:
    diagnostico = get_operacional_diagnostico()
    if diagnostico is None:
        return None
    if diagnostico.cenario == CenarioOperacional.NF_CANCELADA:
        return diagnostico
    decisao = get_operacional_decisao()
    if decisao is None or decisao == DecisaoOperacional.CANCELAR:
        return diagnostico
    if not is_operacional_analise_confirmada():
        return diagnostico
    if not diagnostico.bloqueia_fechamento:
        return diagnostico
    payload = diagnostico.to_dict()
    payload["bloqueia_fechamento"] = False
    return DiagnosticoCarregamento.from_dict(payload)


def cancelar_operacao_pendente() -> None:
    st.session_state.pop("operacional_decisao", None)
    st.session_state.pop(OPERACIONAL_DECISAO_WIDGET_KEY, None)
    st.session_state.pop(OPERACIONAL_ANALISE_CONFIRMADA_KEY, None)
    st.session_state.pop(OPERACIONAL_CONTINUACAO_AUDITORIA_KEY, None)
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
    st.session_state.pop(OPERACIONAL_DECISAO_WIDGET_KEY, None)
    st.session_state.pop(OPERACIONAL_ANALISE_CONFIRMADA_KEY, None)


def get_operacional_decisao() -> DecisaoOperacional | None:
    value = st.session_state.get("operacional_decisao")
    if value:
        try:
            return DecisaoOperacional(str(value))
        except ValueError:
            pass

    widget_value = st.session_state.get(OPERACIONAL_DECISAO_WIDGET_KEY)
    if not widget_value:
        return None
    try:
        decisao = DecisaoOperacional(str(widget_value))
    except ValueError:
        return None
    set_operacional_decisao(decisao)
    return decisao


def _aplicar_decisao_widget(valor: str) -> DecisaoOperacional | None:
    if valor == OPERACIONAL_CONTINUAR_HISTORICO_VALUE:
        diagnostico = get_operacional_diagnostico()
        if diagnostico is None:
            st.session_state.pop("operacional_decisao", None)
            return None
        decisao = inferir_decisao_operacional(diagnostico)
        set_operacional_decisao(decisao)
        _registrar_continuacao_historico(diagnostico)
        return decisao
    try:
        decisao = DecisaoOperacional(str(valor))
    except ValueError:
        st.session_state.pop("operacional_decisao", None)
        return None
    set_operacional_decisao(decisao)
    return decisao


def on_operacional_decisao_widget_change() -> None:
    valor = st.session_state.get(OPERACIONAL_DECISAO_WIDGET_KEY)
    if not valor:
        st.session_state.pop("operacional_decisao", None)
        return
    _aplicar_decisao_widget(str(valor))


def set_operacional_decisao(decisao: DecisaoOperacional) -> None:
    st.session_state["operacional_decisao"] = decisao.value


def confirmar_decisao_operacional_continuacao() -> DecisaoOperacional | None:
    valor = st.session_state.get(OPERACIONAL_DECISAO_WIDGET_KEY)
    if not valor:
        return get_operacional_decisao()
    return _aplicar_decisao_widget(str(valor)) or get_operacional_decisao()


def snapshot_exportacao_documentos(screen_key: str = "minuta") -> dict[str, bool]:
    carregamento_key = f"{screen_key}_pdf_carregamento"
    entrega_key = f"{screen_key}_pdf_entrega"
    xml_key = f"{screen_key}_pdf_xmls"
    return {
        "carregamento_selected": bool(st.session_state.get(carregamento_key, True)),
        "entrega_selected": bool(st.session_state.get(entrega_key, False)),
        "xml_selected": bool(st.session_state.get(xml_key, False)),
    }


def clear_pdf_download_state() -> None:
    st.session_state.pop("pdf_download_payload", None)
    st.session_state.pop("pdf_download_name", None)
    st.session_state.pop("pdf_download_mime", None)


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
        if diagnostico.requer_decisao and get_operacional_decisao() is None:
            if not is_operacional_analise_confirmada():
                return "carregamento_historico"
            return "carregamento_decisao"

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
        clear_pdf_download_state()
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
    elif mode in {"carregamento_historico", "carregamento_decisao"}:
        queue_processing_action({"type": "operacional_cancel"})


def on_baixar_pdf_click() -> None:
    confirmar_decisao_operacional_continuacao()
    decisao = get_operacional_decisao()
    action: dict[str, object] = {"type": "baixar_pdf"}
    if decisao == DecisaoOperacional.REIMPRIMIR:
        action["confirmar_reimpressao"] = True
    clear_pdf_download_state()
    queue_processing_action(action)


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
    diagnostico_efetivo = diagnostico or get_diagnostico_efetivo_fechamento()
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
