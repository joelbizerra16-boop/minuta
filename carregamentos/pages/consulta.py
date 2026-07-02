from __future__ import annotations

from contextlib import contextmanager

import base64
import html
from datetime import date

import pandas as pd
import streamlit as st

from auth.security.session import get_current_user
from carregamentos.bootstrap import get_carregamento_service, get_rastreabilidade_nf_service
from carregamentos.models.carregamento import CarregamentoFiltro
from utils.streamlit_tables import build_consulta_listagem_column_config, build_table_column_config


@contextmanager
def _table_shell(title: str = ""):
    with st.container(border=True):
        if title:
            st.markdown(f"### {html.escape(title)}")
        yield


def _render_pdf_viewer(label: str, pdf_bytes: bytes, download_name: str, key_prefix: str) -> None:
    st.markdown(f"**{html.escape(label)}**")
    if not pdf_bytes:
        st.caption("Documento nao disponivel para este carregamento.")
        return

    col_view, col_download = st.columns([1, 1])
    with col_download:
        st.download_button(
            f"Baixar {label}",
            data=pdf_bytes,
            file_name=download_name,
            mime="application/pdf",
            use_container_width=True,
            key=f"{key_prefix}_download",
        )
    with col_view:
        if st.button(f"Visualizar {label}", use_container_width=True, key=f"{key_prefix}_view"):
            st.session_state[f"{key_prefix}_show"] = not st.session_state.get(f"{key_prefix}_show", False)

    if st.session_state.get(f"{key_prefix}_show", False):
        encoded_pdf = base64.b64encode(pdf_bytes).decode("ascii")
        st.markdown(
            f"""
            <iframe
                src="data:application/pdf;base64,{encoded_pdf}"
                width="100%"
                height="620px"
                style="border: 1px solid rgba(31,58,95,0.12); border-radius: 8px;"
            ></iframe>
            """,
            unsafe_allow_html=True,
        )


def _render_detail_view(carregamento_id: int) -> None:
    service = get_carregamento_service()
    carregamento = service.get_carregamento(carregamento_id)
    if carregamento is None:
        st.error("Carregamento nao encontrado.")
        return

    if st.button("Voltar para consulta", key="consulta_voltar_lista"):
        st.session_state.pop("consulta_carregamento_id", None)
        st.rerun()

    st.markdown("#### Dados Gerais")
    info_col_1, info_col_2, info_col_3 = st.columns(3)
    with info_col_1:
        st.markdown(f"**Carregamento:** {html.escape(carregamento.numero_carregamento)}")
        st.markdown(f"**Data:** {html.escape(carregamento.data)}")
        st.markdown(f"**Hora:** {html.escape(carregamento.hora)}")
    with info_col_2:
        st.markdown(f"**Usuario:** {html.escape(carregamento.usuario)}")
        st.markdown(f"**Motorista:** {html.escape(carregamento.motorista)}")
        st.markdown(f"**Placa:** {html.escape(carregamento.placa)}")
    with info_col_3:
        st.markdown(f"**Filial:** {html.escape(carregamento.filial)}")
        st.markdown(f"**Modalidade:** {html.escape(carregamento.modalidade)}")
        st.markdown(f"**Status:** {html.escape(carregamento.status)}")
        st.markdown(f"**Reentrega:** {'Sim' if carregamento.reentrega else 'Nao'}")

    st.markdown("---")
    st.markdown("#### Resumo")
    resumo_col_1, resumo_col_2, resumo_col_3 = st.columns(3)
    with resumo_col_1:
        st.metric("Total de NFs", carregamento.quantidade_nf)
    with resumo_col_2:
        st.metric("Total de Itens", carregamento.quantidade_itens)
    with resumo_col_3:
        st.metric("Peso Total", f"{carregamento.peso_total / 1000:.3f} t")

    st.markdown("#### NF-es do carregamento")
    if carregamento.itens:
        itens_df = pd.DataFrame(
            [
                {
                    "NF": item.nf,
                    "Produto": item.cprod,
                    "Descricao": item.descricao,
                    "Quantidade": item.quantidade,
                    "Unidade": item.unidade,
                    "Peso": item.peso,
                    "Destinatario": item.destinatario,
                    "Rota": item.rota,
                }
                for item in carregamento.itens
            ]
        )
        st.dataframe(
            itens_df,
            width="stretch",
            hide_index=True,
            column_config=build_table_column_config(itens_df),
        )
    else:
        st.info("Nenhum item registrado para este carregamento.")

    st.markdown("---")
    st.markdown("#### Documentos")

    minuta_bytes = service.read_document(carregamento.minuta_pdf_path)
    romaneio_bytes = service.read_document(carregamento.romaneio_pdf_path)
    doc_key = f"carregamento_{carregamento.id}"

    doc_col_1, doc_col_2 = st.columns(2)
    with doc_col_1:
        _render_pdf_viewer(
            "Minuta de Carregamento",
            minuta_bytes,
            f"minuta_carregamento_{carregamento.numero_carregamento}.pdf",
            f"{doc_key}_minuta",
        )
    with doc_col_2:
        _render_pdf_viewer(
            "Romaneio de Entrega",
            romaneio_bytes,
            f"romaneio_entrega_{carregamento.numero_carregamento}.pdf",
            f"{doc_key}_romaneio",
        )


def render_consulta_carregamentos_page(render_header_callback) -> None:
    render_header_callback(
        "Consulta de NFs Carregadas",
        "Localize rapidamente notas fiscais, produtos e carregamentos finalizados",
    )

    selected_id = st.session_state.get("consulta_carregamento_id")
    if selected_id is not None:
        _render_detail_view(int(selected_id))
        return

    service = get_carregamento_service()
    today = date.today()

    st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
    filter_col_1, filter_col_2, filter_col_3 = st.columns([1.2, 1.2, 0.8])
    with filter_col_1:
        data_inicial = st.date_input("Data inicial", value=today.replace(day=1), key="consulta_data_inicial")
    with filter_col_2:
        data_final = st.date_input("Data final", value=today, key="consulta_data_final")
    with filter_col_3:
        st.markdown("<div style='height: 1.6rem;'></div>", unsafe_allow_html=True)
        if st.button("Consultar", use_container_width=True, key="consulta_buscar"):
            st.session_state["consulta_filtro_aplicado"] = True
            st.session_state["consulta_lista_pagina"] = 0

    termo_pesquisa = st.text_input(
        "Pesquisar",
        placeholder="Digite a NF, produto, descricao, destinatario, rota, motorista, placa ou carregamento...",
        key="consulta_termo_pesquisa",
    )

    if not st.session_state.get("consulta_filtro_aplicado", False):
        st.caption("Informe o periodo e clique em Consultar.")
        return

    filtro = CarregamentoFiltro(
        data_inicial=data_inicial.isoformat(),
        data_final=data_final.isoformat(),
        termo_pesquisa=termo_pesquisa.strip() or None,
    )
    listagem = service.search_itens_listagem(filtro)
    linhas = list(listagem.linhas)
    termo_limpo = termo_pesquisa.strip()

    if termo_limpo:
        relatorio_col_1, relatorio_col_2 = st.columns([4, 1])
        with relatorio_col_2:
            if st.button("Gerar Relatorio", use_container_width=True, key="consulta_gerar_relatorio"):
                try:
                    pdf_bytes = get_rastreabilidade_nf_service().gerar_relatorio_pdf(
                        termo_limpo,
                        get_current_user(),
                    )
                    st.session_state["consulta_rastreabilidade_pdf"] = pdf_bytes
                    st.session_state["consulta_rastreabilidade_nome"] = (
                        f"rastreabilidade_nf_{termo_limpo.replace('/', '-')}.pdf"
                    )
                except ValueError as exc:
                    st.error(str(exc))
        with relatorio_col_1:
            if st.session_state.get("consulta_rastreabilidade_pdf"):
                st.download_button(
                    "Baixar Relatorio",
                    data=st.session_state["consulta_rastreabilidade_pdf"],
                    file_name=st.session_state.get(
                        "consulta_rastreabilidade_nome",
                        "rastreabilidade_nf.pdf",
                    ),
                    mime="application/pdf",
                    use_container_width=True,
                    key="consulta_download_rastreabilidade",
                )

    st.markdown("---")
    resultados_titulo = (
        f"Resultados ({len(linhas)} itens) • "
        f"Carregamentos da NF: {listagem.carregamentos_distintos}"
    )
    with _table_shell(resultados_titulo):
        list_df = pd.DataFrame(
            [
                {
                    "Data": linha.data,
                    "Carregamento": linha.carregamento,
                    "NF": linha.nf,
                    "Produto": linha.produto,
                    "Descricao": linha.descricao,
                    "Quantidade": linha.quantidade,
                    "Peso": linha.peso,
                    "Destinatario": linha.destinatario,
                    "Rota": linha.rota,
                    "Motorista": linha.motorista,
                    "Placa": linha.placa,
                    "Usuario": linha.usuario,
                    "Modalidade": linha.modalidade,
                    "Status": linha.status,
                }
                for linha in linhas
            ]
        )
        if list_df.empty:
            st.info("Nenhum item encontrado para os filtros informados.")
            return

        st.dataframe(
            list_df,
            width="stretch",
            hide_index=True,
            column_config=build_consulta_listagem_column_config(list_df),
        )

    page_size = 50
    total_pages = max((len(linhas) - 1) // page_size + 1, 1)
    current_page = int(st.session_state.get("consulta_lista_pagina", 0))
    current_page = min(max(current_page, 0), total_pages - 1)
    page_col_1, page_col_2, page_col_3 = st.columns([1, 2, 1])
    with page_col_1:
        if st.button("Anterior", disabled=current_page <= 0, key="consulta_pagina_anterior"):
            st.session_state["consulta_lista_pagina"] = current_page - 1
            st.rerun()
    with page_col_2:
        st.caption(f"Pagina {current_page + 1} de {total_pages}")
    with page_col_3:
        if st.button("Proxima", disabled=current_page >= total_pages - 1, key="consulta_pagina_proxima"):
            st.session_state["consulta_lista_pagina"] = current_page + 1
            st.rerun()

    paginated_linhas = linhas[current_page * page_size : (current_page + 1) * page_size]
    with st.expander(f"Acoes ({len(paginated_linhas)})", expanded=False):
        for linha in paginated_linhas:
            action_col_1, action_col_2 = st.columns([5, 1])
            with action_col_1:
                st.caption(
                    f"{linha.data} • NF {linha.nf} • {linha.produto} • Carregamento {linha.carregamento}"
                )
            with action_col_2:
                if st.button(
                    "Visualizar",
                    key=f"consulta_visualizar_{linha.carregamento_id}_{linha.item_index}",
                    use_container_width=True,
                ):
                    st.session_state["consulta_carregamento_id"] = linha.carregamento_id
                    st.rerun()
