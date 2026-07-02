from __future__ import annotations

import hashlib
import html
from datetime import datetime, timezone
from typing import Any

import pandas as pd
import streamlit as st

from carregamentos.bootstrap import get_analise_operacional_service
from carregamentos.integration import get_operacional_diagnostico
from carregamentos.models.auditoria_nf import (
    AuditoriaNfLote,
    NfAuditoriaCard,
    NfAuditoriaEvento,
    TipoOperacaoNf,
)
from carregamentos.models.carregamento import normalize_chave_nfe, normalize_nf_number
from carregamentos.repository.sql_auditoria_nf_repository import (
    MovimentacaoNfExtratoRegistro,
    SqlAuditoriaNfRepository,
)
from infrastructure.models.constants import (
    AUDIT_EVENTO_COMPLEMENTACAO,
    AUDIT_EVENTO_ENTREGA_BALCAO,
    AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
    AUDIT_EVENTO_REENTREGA,
    AUDIT_EVENTO_REIMPRESSAO,
    HISTORICO_EVENTO_COMPLEMENTACAO,
    HISTORICO_EVENTO_ENTREGA_BALCAO,
    HISTORICO_EVENTO_FINALIZACAO,
    HISTORICO_EVENTO_REENTREGA,
)
from utils.streamlit_tables import build_auditoria_nf_expansion_column_config

# Cada linha: tuple de pares (coluna, valor) — hashavel para @st.cache_data.
ProcessedDfPayload = tuple[tuple[tuple[str, Any], ...], ...]

_MAIN_COL_WEIGHTS = [0.08, 0.22, 0.10, 0.12, 0.08, 0.06, 0.10, 0.03]
_MAIN_HEADERS = [
    "NF",
    "Cliente",
    "Situacao",
    "Ultima Operacao",
    "Data",
    "Hora",
    "Usuario",
    "",
]

_OPERACAO_DISPLAY: dict[TipoOperacaoNf, str] = {
    TipoOperacaoNf.IMPRESSAO_ORIGINAL: "Primeira Impressao",
    TipoOperacaoNf.REIMPRESSAO: "Reimpressao",
    TipoOperacaoNf.COMPLEMENTACAO: "Complementacao",
    TipoOperacaoNf.REENTREGA: "Reentrega",
    TipoOperacaoNf.CANCELAMENTO: "Cancelamento",
}

_OPERACAO_STYLE: dict[str, tuple[str, str]] = {
    "Primeira Impressao": ("#2563eb", "#ffffff"),
    "Carregamento": ("#2563eb", "#ffffff"),
    "Reimpressao": ("#ea580c", "#ffffff"),
    "Complementacao": ("#16a34a", "#ffffff"),
    "Reentrega": ("#7c3aed", "#ffffff"),
    "Cancelamento": ("#dc2626", "#ffffff"),
    "Nova": ("#64748b", "#ffffff"),
}

_ETAPA_OBSERVACAO_PADRAO: dict[str, str] = {
    "Carregamento": "Carregamento realizado",
    "Reimpressao": "Reimpressao da minuta",
    "Complementacao": "Complementacao da carga",
    "Reentrega": "Nova entrega",
    "Cancelamento": "Cancelamento operacional",
}

_EXPANSION_COLUMNS = [
    "Etapa",
    "Veiculo",
    "Placa",
    "Motorista",
    "Data",
    "Hora",
    "Usuario",
    "Carregamento",
    "IdCarga",
    "Rota",
    "Observacao",
]


def _json_safe_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "item"):
        return value.item()
    return str(value)


def _serialize_processed_df(processed_df: pd.DataFrame) -> ProcessedDfPayload:
    records: list[dict[str, Any]] = processed_df.to_dict(orient="records")
    return tuple(
        tuple((str(key), _json_safe_value(val)) for key, val in row.items())
        for row in records
    )


def _deserialize_processed_df(payload: ProcessedDfPayload) -> pd.DataFrame:
    if not payload:
        return pd.DataFrame()
    records: list[dict[str, Any]] = [
        {key: value for key, value in row}
        for row in payload
    ]
    return pd.DataFrame.from_records(records)


def _cache_key_from_payload(payload: ProcessedDfPayload) -> str:
    digest = hashlib.sha256(repr(payload).encode("utf-8")).hexdigest()
    return digest[:16]


def _expanded_state_key(cache_key: str, token: str, scope: str = "expander") -> str:
    return f"auditoria_nf_expanded_{scope}_{cache_key}_{token}"


def _clear_other_expanded_rows(cache_key: str, keep_token: str, scope: str = "expander") -> None:
    prefix = f"auditoria_nf_expanded_{scope}_{cache_key}_"
    keep_key = _expanded_state_key(cache_key, keep_token, scope)
    for key in list(st.session_state.keys()):
        if str(key).startswith(prefix) and key != keep_key:
            st.session_state.pop(key, None)


@st.cache_data(show_spinner=False, ttl=300)
def _carregar_auditoria_nf_cache(cache_key: str, processed_df_payload: ProcessedDfPayload) -> dict:
    if not processed_df_payload:
        return {}
    processed_df = _deserialize_processed_df(processed_df_payload)
    auditoria = get_analise_operacional_service().montar_auditoria_nfs_lote(processed_df)
    return auditoria.to_dict()


def _lista_session_key(cache_key: str, scope: str) -> str:
    return f"auditoria_nf_lista_{scope}_{cache_key}"


def _clear_expanded_rows(cache_key: str, scope: str) -> None:
    prefix = f"auditoria_nf_expanded_{scope}_{cache_key}_"
    for key in list(st.session_state.keys()):
        if str(key).startswith(prefix):
            st.session_state.pop(key, None)


def render_auditoria_nf_expander(*, processed_df) -> None:
    if processed_df is None or processed_df.empty:
        return

    payload = _serialize_processed_df(processed_df)
    cache_key = _cache_key_from_payload(payload)
    session_token = f"auditoria_nf_loaded_{cache_key}"

    with st.expander("Consultar historico operacional das Notas Fiscais", expanded=False):
        if not st.session_state.get(session_token):
            if st.button(
                "Carregar historico das NFs",
                key=f"btn_{session_token}",
                use_container_width=True,
                type="secondary",
            ):
                st.session_state[session_token] = True
                st.rerun()
            st.caption("Consulta opcional. A listagem so e carregada ao solicitar.")
            return

        with st.spinner("Carregando historico operacional das NFs..."):
            data = _carregar_auditoria_nf_cache(cache_key, payload)
        if not data:
            st.warning("Nao foi possivel carregar o historico das notas fiscais.")
            return

        auditoria = AuditoriaNfLote.from_dict(data)
        st.markdown('<div class="nf-historico-compacto">', unsafe_allow_html=True)
        _render_nf_listview(auditoria, cache_key=cache_key, scope="expander")
        st.markdown("</div>", unsafe_allow_html=True)


def render_historico_nfs_contexto(processed_df, *, scope: str = "painel") -> bool:
    """Resumo simples e listagem sob demanda para apoio a decisao operacional."""
    if processed_df is None or processed_df.empty:
        return False

    diagnostico = get_operacional_diagnostico()
    if diagnostico is None or diagnostico.nfs_existentes <= 0:
        return False

    payload = _serialize_processed_df(processed_df)
    cache_key = _cache_key_from_payload(payload)
    lista_key = _lista_session_key(cache_key, scope)

    st.markdown('<div class="nf-historico-compacto">', unsafe_allow_html=True)
    _render_resumo_simples(diagnostico)
    _render_acoes_historico(cache_key=cache_key, scope=scope, lista_key=lista_key)

    if st.session_state.get(lista_key):
        with st.spinner("Carregando historico das notas fiscais..."):
            data = _carregar_auditoria_nf_cache(cache_key, payload)
        if not data:
            st.warning("Nao foi possivel carregar o historico das notas fiscais.")
        else:
            auditoria = AuditoriaNfLote.from_dict(data)
            _render_nf_listview(auditoria, cache_key=cache_key, scope=scope)

    st.markdown("</div>", unsafe_allow_html=True)
    return True


def _render_resumo_simples(diagnostico) -> None:
    st.markdown(
        '<p class="nf-historico-mensagem">Foram encontradas NFs com historico operacional.</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="nf-historico-resumo-linhas">
            <div class="nf-historico-resumo-linha">
                <span>Total de NFs</span><strong>{diagnostico.nfs_total}</strong>
            </div>
            <div class="nf-historico-resumo-linha">
                <span>NFs novas</span><strong>{diagnostico.nfs_novas}</strong>
            </div>
            <div class="nf-historico-resumo-linha">
                <span>NFs com historico</span><strong>{diagnostico.nfs_existentes}</strong>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_acoes_historico(*, cache_key: str, scope: str, lista_key: str) -> None:
    col_continuar, col_visualizar = st.columns(2, gap="small")
    with col_continuar:
        if st.button(
            "Continuar processamento",
            key=f"btn_nf_continuar_{scope}_{cache_key}",
            use_container_width=True,
            type="primary",
        ):
            st.session_state.pop(lista_key, None)
            _clear_expanded_rows(cache_key, scope)
            st.session_state[f"auditoria_nf_foco_decisao_{scope}"] = True
            st.rerun()
    with col_visualizar:
        if st.button(
            "Visualizar historico",
            key=f"btn_nf_visualizar_{scope}_{cache_key}",
            use_container_width=True,
            type="secondary",
        ):
            st.session_state[lista_key] = True
            st.rerun()


def _card_tem_historico(card: NfAuditoriaCard) -> bool:
    if card.quantidade_utilizacoes > 0 or card.eventos:
        return True
    return "Ja utilizada" in card.situacao_atual


def _render_cliente_cell(cliente: str) -> None:
    texto = str(cliente or "--").strip() or "--"
    safe = html.escape(texto)
    st.markdown(
        f'<span class="nf-cliente-cell" title="{safe}">{safe}</span>',
        unsafe_allow_html=True,
    )


def _render_situacao_cell(card: NfAuditoriaCard) -> None:
    label = _situacao_display(card)
    if label in {"Ja utilizada", "Reentrega"}:
        st.markdown(
            f'<span class="nf-badge-utilizada">{html.escape(label)}</span>',
            unsafe_allow_html=True,
        )
        return
    st.write(label)


def _render_nf_listview(auditoria: AuditoriaNfLote, *, cache_key: str, scope: str = "expander") -> None:
    cards = sorted(
        auditoria.cards,
        key=lambda item: (0 if _card_tem_historico(item) else 1, item.nf),
    )
    _render_listview_header()
    for index, card in enumerate(cards):
        _render_listview_row(card, cache_key=cache_key, row_index=index, scope=scope)


def _render_listview_header() -> None:
    header_cols = st.columns(_MAIN_COL_WEIGHTS, gap="small")
    for col, label in zip(header_cols, _MAIN_HEADERS):
        with col:
            if label:
                st.markdown(f"**{label}**")


def _render_listview_row(
    card: NfAuditoriaCard,
    *,
    cache_key: str,
    row_index: int,
    scope: str = "expander",
) -> None:
    state_key = _expanded_state_key(cache_key, card.token, scope)
    expanded = bool(st.session_state.get(state_key, False))
    ultima_operacao, _, ultima_data, ultima_hora, ultimo_usuario = _resumo_linha_principal(card)
    row_class = "nf-row-utilizada" if _card_tem_historico(card) else "nf-row-nova"
    st.markdown(f'<div class="{row_class} nf-listview-row">', unsafe_allow_html=True)

    row_cols = st.columns(_MAIN_COL_WEIGHTS, gap="small")
    with row_cols[0]:
        st.write(card.nf)
    with row_cols[1]:
        _render_cliente_cell(card.cliente)
    with row_cols[2]:
        _render_situacao_cell(card)
    with row_cols[3]:
        st.write(ultima_operacao)
    with row_cols[4]:
        st.write(ultima_data)
    with row_cols[5]:
        st.write(ultima_hora)
    with row_cols[6]:
        st.write(ultimo_usuario)
    with row_cols[7]:
        toggle_label = "▼" if expanded else "▶"
        if st.button(
            toggle_label,
            key=f"btn_{scope}_{state_key}_{row_index}",
            help=f"Historico da NF {card.nf}",
        ):
            if expanded:
                st.session_state.pop(state_key, None)
            else:
                _clear_other_expanded_rows(cache_key, card.token, scope)
                st.session_state[state_key] = True
            st.rerun()

    if expanded:
        _render_history_listview(card, cache_key=cache_key)

    st.markdown("</div>", unsafe_allow_html=True)


def _resolver_identificadores_nf(card: NfAuditoriaCard) -> tuple[str, str]:
    chave = normalize_chave_nfe(card.token)
    numero = normalize_nf_number(card.nf) or str(card.nf or "").strip()
    return numero, chave


@st.cache_data(show_spinner=False, ttl=300)
def _carregar_extrato_nf_cache(cache_key: str, numero_nf: str, chave_nfe: str) -> tuple[dict[str, Any], ...]:
    movimentacoes = SqlAuditoriaNfRepository().buscar_extrato_movimentacoes_nf(
        numero_nf=numero_nf,
        chave_nfe=chave_nfe,
    )
    return tuple(
        {
            "fonte": item.fonte,
            "evento_id": item.evento_id,
            "carregamento_id": item.carregamento_id,
            "evento": item.evento,
            "criado_em": item.criado_em.isoformat() if item.criado_em is not None else "",
            "descricao": item.descricao,
            "metadados_json": item.metadados_json,
            "numero_carregamento": item.numero_carregamento,
            "motorista": item.motorista,
            "placa": item.placa,
            "modalidade": item.modalidade,
            "status": item.status,
            "usuario": item.usuario,
            "rota": item.rota,
            "destinatario": item.destinatario,
        }
        for item in movimentacoes
    )


def _restaurar_movimentacoes(payload: tuple[dict[str, Any], ...]) -> list[MovimentacaoNfExtratoRegistro]:
    registros: list[MovimentacaoNfExtratoRegistro] = []
    for item in payload:
        criado_em_raw = str(item.get("criado_em", "") or "").strip()
        criado_em = None
        if criado_em_raw:
            try:
                criado_em = datetime.fromisoformat(criado_em_raw)
            except ValueError:
                criado_em = None
        registros.append(
            MovimentacaoNfExtratoRegistro(
                fonte=str(item.get("fonte", "") or ""),
                evento_id=int(item.get("evento_id", 0) or 0),
                carregamento_id=int(item.get("carregamento_id", 0) or 0),
                evento=str(item.get("evento", "") or ""),
                criado_em=criado_em,
                descricao=str(item.get("descricao", "") or ""),
                metadados_json=item.get("metadados_json"),
                numero_carregamento=str(item.get("numero_carregamento", "") or ""),
                motorista=str(item.get("motorista", "") or ""),
                placa=str(item.get("placa", "") or ""),
                modalidade=str(item.get("modalidade", "") or ""),
                status=str(item.get("status", "") or ""),
                usuario=str(item.get("usuario", "") or ""),
                rota=str(item.get("rota", "") or ""),
                destinatario=str(item.get("destinatario", "") or ""),
            )
        )
    return registros


def _classificar_etapa(evento: str) -> str | None:
    evento_norm = str(evento or "").strip().upper()
    if evento_norm in {
        AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
        HISTORICO_EVENTO_FINALIZACAO,
        HISTORICO_EVENTO_ENTREGA_BALCAO,
        AUDIT_EVENTO_ENTREGA_BALCAO,
    }:
        return "Carregamento"
    if evento_norm == AUDIT_EVENTO_REIMPRESSAO:
        return "Reimpressao"
    if evento_norm in {AUDIT_EVENTO_COMPLEMENTACAO, HISTORICO_EVENTO_COMPLEMENTACAO}:
        return "Complementacao"
    if evento_norm in {AUDIT_EVENTO_REENTREGA, HISTORICO_EVENTO_REENTREGA}:
        return "Reentrega"
    return None


def _veiculo_display(modalidade: str) -> str:
    texto = str(modalidade or "").strip()
    if not texto:
        return "--"
    upper = texto.upper()
    if upper in {"VEICULO", "VEÍCULO", "ROTA"}:
        return "Caminhao"
    if "BALCAO" in upper or "BALCÃO" in upper:
        return "Balcao"
    return texto


def _formatar_data_hora_movimentacao(
    criado_em: datetime | None,
    *,
    carregamento_id: int,
) -> tuple[str, str, float]:
    if criado_em is not None:
        dt = criado_em
        if isinstance(dt, datetime) and dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        local = dt.astimezone()
        return local.strftime("%d/%m/%Y"), local.strftime("%H:%M"), local.timestamp()
    return "--", "--", float(carregamento_id)


def _observacao_movimentacao(etapa: str, descricao: str) -> str:
    texto = str(descricao or "").strip()
    if texto:
        return texto
    return _ETAPA_OBSERVACAO_PADRAO.get(etapa, etapa)


def _filtrar_movimentacoes_expansao(
    movimentacoes: list[MovimentacaoNfExtratoRegistro],
) -> list[MovimentacaoNfExtratoRegistro]:
    audit_primeira_por_carregamento = {
        int(item.carregamento_id)
        for item in movimentacoes
        if item.fonte == "auditoria" and str(item.evento or "") == AUDIT_EVENTO_PRIMEIRA_IMPRESSAO
    }
    filtradas: list[MovimentacaoNfExtratoRegistro] = []
    for item in movimentacoes:
        if (
            item.fonte == "historico"
            and str(item.evento or "") == HISTORICO_EVENTO_FINALIZACAO
            and int(item.carregamento_id) in audit_primeira_por_carregamento
        ):
            continue
        if _classificar_etapa(item.evento) is None:
            continue
        filtradas.append(item)
    return filtradas


def build_expansion_dataframe(movimentacoes: list[MovimentacaoNfExtratoRegistro]) -> pd.DataFrame:
    if not movimentacoes:
        return pd.DataFrame(columns=_EXPANSION_COLUMNS)

    linhas: list[dict[str, str]] = []
    for item in _filtrar_movimentacoes_expansao(movimentacoes):
        etapa = _classificar_etapa(item.evento) or "--"
        data, hora, _ = _formatar_data_hora_movimentacao(
            item.criado_em,
            carregamento_id=item.carregamento_id,
        )
        sem_veiculo = etapa == "Reimpressao"
        linhas.append(
            {
                "Etapa": etapa,
                "Veiculo": "--" if sem_veiculo else _veiculo_display(item.modalidade),
                "Placa": "--" if sem_veiculo else (item.placa or "--"),
                "Motorista": "--" if sem_veiculo else (item.motorista or "--"),
                "Data": data,
                "Hora": hora,
                "Usuario": item.usuario or "--",
                "Carregamento": item.numero_carregamento or "--",
                "IdCarga": str(item.carregamento_id or "--"),
                "Rota": item.rota or "--",
                "Observacao": _observacao_movimentacao(etapa, item.descricao),
            }
        )

    if not linhas:
        return pd.DataFrame(columns=_EXPANSION_COLUMNS)

    dataframe = pd.DataFrame(linhas)
    dataframe["_ordenacao"] = pd.to_datetime(
        dataframe["Data"] + " " + dataframe["Hora"],
        format="%d/%m/%Y %H:%M",
        errors="coerce",
    )
    dataframe = dataframe.sort_values(by=["_ordenacao", "IdCarga"], ascending=[True, True], na_position="last")
    return dataframe.drop(columns=["_ordenacao"]).reset_index(drop=True)


def build_history_dataframe(card: NfAuditoriaCard) -> pd.DataFrame:
    """Compatibilidade de testes: monta extrato a partir da consulta dedicada da expansao."""
    if card.eventos:
        return build_expansion_dataframe(_eventos_card_para_movimentacoes(card))
    numero_nf, chave_nfe = _resolver_identificadores_nf(card)
    payload = _carregar_extrato_nf_cache(f"nf:{numero_nf}:{chave_nfe}", numero_nf, chave_nfe)
    return build_expansion_dataframe(_restaurar_movimentacoes(payload))


def _style_etapa_cell(value: object) -> str:
    label = str(value or "").strip()
    background, foreground = _OPERACAO_STYLE.get(label, ("#2563eb", "#ffffff"))
    return "; ".join(
        [
            f"background-color: {background}",
            f"color: {foreground}",
            "font-weight: 700",
            "text-align: center",
            "border-radius: 999px",
        ]
    )


def _style_expansion_dataframe(dataframe: pd.DataFrame):
    if dataframe.empty:
        return dataframe
    return dataframe.style.map(_style_etapa_cell, subset=["Etapa"]).set_properties(
        subset=["Etapa"],
        **{"text-align": "center"},
    )


def _carregar_extrato_expansao(card: NfAuditoriaCard, *, cache_key: str) -> pd.DataFrame:
    numero_nf, chave_nfe = _resolver_identificadores_nf(card)
    payload = _carregar_extrato_nf_cache(f"{cache_key}:{card.token}", numero_nf, chave_nfe)
    movimentacoes = _restaurar_movimentacoes(payload)
    if not movimentacoes and card.eventos:
        movimentacoes = _eventos_card_para_movimentacoes(card)
    return build_expansion_dataframe(movimentacoes)


def _render_history_listview(card: NfAuditoriaCard, *, cache_key: str) -> None:
    st.markdown('<div class="nf-extrato-operacional">', unsafe_allow_html=True)
    history_df = _carregar_extrato_expansao(card, cache_key=cache_key)
    if history_df.empty:
        st.info("Nenhuma movimentacao operacional registrada para esta nota fiscal.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    styled_df = _style_expansion_dataframe(history_df)
    st.dataframe(
        styled_df,
        width="stretch",
        hide_index=True,
        column_config=build_auditoria_nf_expansion_column_config(history_df),
    )
    st.markdown("</div>", unsafe_allow_html=True)


def _eventos_card_para_movimentacoes(card: NfAuditoriaCard) -> list[MovimentacaoNfExtratoRegistro]:
    fallback: list[MovimentacaoNfExtratoRegistro] = []
    for evento in sorted(card.eventos, key=lambda item: item.ordenacao):
        etapa = _OPERACAO_DISPLAY.get(evento.operacao, evento.operacao_label)
        evento_codigo = {
            "Primeira Impressao": AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
            "Carregamento": HISTORICO_EVENTO_FINALIZACAO,
            "Reimpressao": AUDIT_EVENTO_REIMPRESSAO,
            "Complementacao": AUDIT_EVENTO_COMPLEMENTACAO,
            "Reentrega": AUDIT_EVENTO_REENTREGA,
        }.get(etapa, AUDIT_EVENTO_PRIMEIRA_IMPRESSAO)
        try:
            ordenacao_dt = datetime.strptime(f"{evento.data} {evento.hora}", "%d/%m/%Y %H:%M")
            criado_em = ordenacao_dt.replace(tzinfo=timezone.utc)
        except ValueError:
            criado_em = None
        fallback.append(
            MovimentacaoNfExtratoRegistro(
                fonte="auditoria",
                evento_id=int(evento.ordenacao),
                carregamento_id=int(evento.ordenacao),
                evento=evento_codigo,
                criado_em=criado_em,
                descricao=_ETAPA_OBSERVACAO_PADRAO.get(etapa, etapa),
                metadados_json=None,
                numero_carregamento=evento.numero_carregamento,
                motorista=evento.motorista,
                placa=evento.placa,
                modalidade=evento.tipo_operacao,
                status=evento.status_carregamento,
                usuario=evento.usuario,
                rota=evento.rota,
                destinatario=card.cliente,
            )
        )
    return fallback


def _situacao_display(card: NfAuditoriaCard) -> str:
    if card.eventos and card.eventos[0].operacao == TipoOperacaoNf.REENTREGA:
        return "Reentrega"
    if "Nunca utilizada" in card.situacao_atual:
        return "Nova"
    if "Ja utilizada" in card.situacao_atual:
        return "Ja utilizada"
    return card.situacao_atual


def _resumo_linha_principal(card: NfAuditoriaCard) -> tuple[str, str, str, str, str]:
    if not card.eventos:
        return "Nova", "--", "--", "--", "--"
    evento = card.eventos[0]
    return (
        _operacao_display(evento),
        str(evento.numero_carregamento or "--"),
        str(evento.data or "--"),
        str(evento.hora or "--"),
        str(evento.usuario or "--"),
    )


def _operacao_display(evento: NfAuditoriaEvento) -> str:
    return _OPERACAO_DISPLAY.get(evento.operacao, evento.operacao_label)
