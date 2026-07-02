from __future__ import annotations

import streamlit as st
import pandas as pd

from carregamentos.bootstrap import get_historico_carregamento_service
from carregamentos.models.historico_carregamento_painel import HistoricoCarregamentoPainel
from utils.streamlit_tables import build_table_column_config


def _format_brl(value: float) -> str:
    texto = f"{float(value):,.2f}"
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _format_kg(value: float) -> str:
    return f"{float(value):,.3f} kg".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False, ttl=300)
def _carregar_painel_cache(carregamento_id: int, excel_contexto: str) -> dict:
    painel = get_historico_carregamento_service().montar_painel_auditoria(
        carregamento_id,
        excel_contexto=excel_contexto,
    )
    if painel is None:
        return {}
    return painel.to_dict()


def render_historico_operacional_expander(
    *,
    carregamento_id: int | None,
    excel_contexto: str = "",
) -> None:
    if not carregamento_id:
        return

    cache_token = f"historico_carregamento_loaded_{carregamento_id}"
    with st.expander("Historico Operacional do Carregamento", expanded=False):
        if not st.session_state.get(cache_token):
            if st.button(
                "Consultar historico completo",
                key=f"btn_{cache_token}",
                use_container_width=True,
                type="secondary",
            ):
                st.session_state[cache_token] = True
                st.rerun()
            st.caption(
                "O historico e carregado somente quando solicitado, sem impactar o processamento da minuta."
            )
            return

        payload = _carregar_painel_cache(int(carregamento_id), str(excel_contexto or ""))
        if not payload:
            st.warning("Nao foi possivel carregar o historico deste carregamento.")
            return

        painel = HistoricoCarregamentoPainel.from_dict(payload)
        _render_estatisticas(painel)
        _render_tabela_nfs(painel)
        _render_historico_impressoes(painel)
        if painel.complementacoes:
            _render_complementacoes(painel)
        if painel.reentregas:
            _render_reentregas(painel)


def _render_estatisticas(painel: HistoricoCarregamentoPainel) -> None:
    stats = painel.estatisticas
    st.markdown("**Indicadores do carregamento**")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Carregamento", stats.numero_carregamento)
        st.metric("Total de NFs", stats.total_nfs)
    with col2:
        st.metric("Peso total", _format_kg(stats.peso_total_kg))
        st.metric("Valor total", _format_brl(stats.valor_total))
    with col3:
        st.metric(
            "Primeira impressao",
            f"{stats.primeira_impressao_data} {stats.primeira_impressao_hora}".strip(),
        )
        st.metric(
            "Ultima impressao",
            f"{stats.ultima_impressao_data} {stats.ultima_impressao_hora}".strip(),
        )

    col4, col5, col6 = st.columns(3)
    with col4:
        st.metric("Reimpressoes", stats.quantidade_reimpressoes)
    with col5:
        st.metric("Complementacoes", stats.quantidade_complementacoes)
    with col6:
        st.metric("Reentregas", stats.quantidade_reentregas)

    st.caption(
        f"Consulta realizada em {painel.data_analise} • "
        f"Excel em uso: {painel.excel_contexto}"
    )
    st.markdown("---")


def _render_tabela_nfs(painel: HistoricoCarregamentoPainel) -> None:
    st.markdown("**Notas fiscais do carregamento**")
    if not painel.nfs:
        st.info("Nenhuma NF encontrada para este carregamento.")
        return

    rows = []
    for item in painel.nfs:
        rows.append(
            {
                "NF": item.nf,
                "Cliente": item.cliente,
                "Cidade": item.cidade,
                "UF": item.uf,
                "Peso": _format_kg(item.peso_kg),
                "Valor NF": _format_brl(item.valor_nf),
                "Primeira utilizacao": f"{item.primeira_utilizacao_data} {item.primeira_utilizacao_hora}",
                "Usuario": item.usuario_carregamento,
                "Carregamento": item.numero_carregamento,
                "Reimpressoes": item.quantidade_reimpressoes,
                "Ultima reimpressao": f"{item.ultima_reimpressao_data} {item.ultima_reimpressao_hora}".strip(),
                "Ultimo usuario": item.ultimo_usuario_impressao,
                "Status": item.status_atual,
                "Origem": item.origem,
                "Excel": item.excel_utilizado,
                "XML": item.xml_status,
                "Arquivo XML": item.xml_arquivo,
            }
        )

    df = pd.DataFrame(rows)
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config=build_table_column_config(df),
        row_height=48,
    )
    st.markdown("---")


def _render_historico_impressoes(painel: HistoricoCarregamentoPainel) -> None:
    with st.expander("Historico de Impressoes", expanded=False):
        if not painel.impressoes:
            st.info("Nenhuma impressao registrada para este carregamento.")
            return
        df = pd.DataFrame(
            [
                {
                    "Data": item.data,
                    "Hora": item.hora,
                    "Usuario": item.usuario,
                    "Tipo": item.tipo,
                    "Resultado": item.resultado,
                }
                for item in painel.impressoes
            ]
        )
        st.dataframe(df, use_container_width=True, hide_index=True)


def _render_complementacoes(painel: HistoricoCarregamentoPainel) -> None:
    with st.expander("Historico de Complementacoes", expanded=False):
        df = pd.DataFrame(
            [
                {
                    "Data": item.data,
                    "Hora": item.hora,
                    "Usuario": item.usuario,
                    "NFs adicionadas": item.nfs_adicionadas,
                    "Observacao": item.observacao,
                }
                for item in painel.complementacoes
            ]
        )
        st.dataframe(df, use_container_width=True, hide_index=True)


def _render_reentregas(painel: HistoricoCarregamentoPainel) -> None:
    with st.expander("Historico de Reentregas", expanded=False):
        df = pd.DataFrame(
            [
                {
                    "Data": item.data,
                    "Hora": item.hora,
                    "Usuario": item.usuario,
                    "Motivo": item.motivo,
                    "Status": item.status,
                }
                for item in painel.reentregas
            ]
        )
        st.dataframe(df, use_container_width=True, hide_index=True)
