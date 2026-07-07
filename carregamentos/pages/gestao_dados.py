from __future__ import annotations

import html
from datetime import date

import streamlit as st

from auth.security.session import get_current_user, is_admin
from carregamentos.bootstrap import (
    get_execucao_retencao_service,
    get_gestao_capacidade_service,
    get_gestao_dados_service,
    get_simulacao_retencao_service,
)
from carregamentos.models.capacidade import CapacidadeOperacional, FaixaCapacidade, PreviaRetencaoCapacidade
from carregamentos.models.execucao_retencao import ConfirmacaoRetencao, ResultadoRetencao
from carregamentos.models.retencao import GestaoDadosPainel
from carregamentos.models.simulacao_retencao import PacoteRetencaoUnitario, RelatorioSimulacaoRetencao, SaudePacote
from carregamentos.services.execucao_retencao_service import RetencaoExecucaoError
from carregamentos.services.gestao_capacidade_service import GestaoCapacidadeError
from core.retention_policy import CAPACITY_ORANGE_MIN_PERCENT
from infrastructure.services.database_usage_service import UsoBancoDados

_SESSION_SIMULACAO = "gestao_dados_simulacao_relatorio"
_SESSION_CONFIRMACAO = "gestao_dados_retencao_confirmacao"
_SESSION_RESULTADO = "gestao_dados_retencao_resultado"
_SESSION_CAPACIDADE_PREVIA = "gestao_dados_capacidade_previa"
_SESSION_CAPACIDADE_CONFIRMACAO = "gestao_dados_capacidade_confirmacao"
_SESSION_CAPACIDADE_INICIAR = "gestao_dados_capacidade_iniciar"
_STYLES_INJECTED = "_gestao_dados_styles_v3_injected"

CardSpec = tuple[str, str, str]


def _formatar_data_br(value: date) -> str:
    return value.strftime("%d/%m/%Y")


def _formatar_bytes(value: int | None, *, indisponivel: str = "Indisponivel") -> str:
    if value is None:
        return indisponivel
    total = max(int(value), 0)
    if total >= 1024 * 1024 * 1024:
        return f"{total / (1024 * 1024 * 1024):.2f} GB"
    if total >= 1024 * 1024:
        return f"{total / (1024 * 1024):.1f} MB"
    if total >= 1024:
        return f"{total / 1024:.1f} KB"
    return f"{total} B"


def _inject_gestao_dados_styles() -> None:
    if st.session_state.get(_STYLES_INJECTED):
        return
    st.session_state[_STYLES_INJECTED] = True
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

.gd-page-root,
.gd-page-root .gd-dash-section,
.gd-page-root .gd-dash-card,
.gd-page-root .gd-dash-footnote {
    font-family: "Inter", "Source Sans 3", "IBM Plex Sans", "Segoe UI Variable",
        "Segoe UI", "Noto Sans", system-ui, -apple-system, sans-serif;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
}

.gd-dash-section {
    color: #1E293B;
    font-size: 18px;
    font-weight: 500;
    letter-spacing: 0.2px;
    line-height: 1.15;
    margin: 0.65rem 0 0.55rem 0;
    padding: 0;
}
.gd-dash-card {
    background: #FFFFFF;
    border: 1px solid #1F3A5F;
    border-radius: 10px;
    box-shadow: 0 2px 8px rgba(31, 58, 95, 0.06);
    padding: 0.55rem 0.65rem 0.6rem;
    min-height: 74px;
    height: 100%;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 0.12rem;
}
.gd-dash-card-compact {
    min-height: 66px;
    padding: 0.45rem 0.55rem 0.5rem;
}
.gd-dash-label {
    color: #64748B;
    font-size: 12px;
    font-weight: 400;
    letter-spacing: 0.3px;
    line-height: 1.12;
    white-space: normal;
}
.gd-dash-value {
    color: #1E293B;
    font-size: 28px;
    font-weight: 500;
    letter-spacing: 0;
    line-height: 1.05;
    white-space: normal;
    overflow-wrap: anywhere;
    word-break: break-word;
    font-variant-numeric: tabular-nums;
}
.gd-dash-card-compact .gd-dash-value {
    font-size: 26px;
}
.gd-dash-value-text {
    font-size: 15px;
    font-weight: 500;
    letter-spacing: 0;
    line-height: 1.18;
    color: #334155;
}
.gd-dash-note {
    color: #94A3B8;
    font-size: 11px;
    font-weight: 400;
    letter-spacing: 0;
    line-height: 1.15;
    margin-top: 0.08rem;
    white-space: normal;
    overflow-wrap: anywhere;
}
.gd-dash-footnote {
    color: #94A3B8;
    font-size: 11px;
    font-weight: 400;
    letter-spacing: 0;
    line-height: 1.2;
    margin: 0.35rem 0 0.15rem;
}
.gd-page-root [data-testid="stCaptionContainer"] p,
.gd-page-root [data-testid="stCaptionContainer"] {
    color: #94A3B8 !important;
    font-family: "Inter", "Source Sans 3", "IBM Plex Sans", "Segoe UI Variable",
        "Segoe UI", "Noto Sans", system-ui, sans-serif !important;
    font-size: 11px !important;
    font-weight: 400 !important;
    line-height: 1.2 !important;
}
[data-testid="column"] .gd-dash-card {
    width: 100%;
}
.gd-cap-banner {
    border: 1px solid;
    border-radius: 10px;
    padding: 0.65rem 0.85rem;
    margin: 0 0 0.75rem 0;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    background: #FFFFFF;
}
.gd-cap-banner-title {
    font-size: 13px;
    font-weight: 500;
    color: #64748B;
    letter-spacing: 0.2px;
}
.gd-cap-banner-status {
    font-size: 15px;
    font-weight: 600;
    letter-spacing: 0.1px;
}
.gd-cap-bar-wrap {
    display: flex;
    align-items: center;
    gap: 0.55rem;
    font-family: "Inter", monospace, sans-serif;
    font-size: 15px;
    font-weight: 500;
    color: #1E293B;
    font-variant-numeric: tabular-nums;
}
.gd-cap-bar {
    letter-spacing: 1px;
    line-height: 1;
}
</style>
        """,
        unsafe_allow_html=True,
    )


def _dash_card_markup(label: str, value: str, *, note: str = "", compact: bool = False) -> str:
    value_class = "gd-dash-value gd-dash-value-text" if len(value) > 24 else "gd-dash-value"
    card_class = "gd-dash-card gd-dash-card-compact" if compact else "gd-dash-card"
    note_html = f'<div class="gd-dash-note">{html.escape(note)}</div>' if note else ""
    return (
        f'<div class="{card_class}">'
        f'<div class="gd-dash-label">{html.escape(label)}</div>'
        f'<div class="{value_class}">{html.escape(value)}</div>'
        f"{note_html}"
        f"</div>"
    )


def _render_section(title: str) -> None:
    st.markdown(f'<div class="gd-dash-section">{html.escape(title)}</div>', unsafe_allow_html=True)


def _render_cards_row(cards: list[CardSpec], *, compact: bool = False) -> None:
    if not cards:
        return
    columns = st.columns(len(cards), gap="small")
    for column, (label, value, note) in zip(columns, cards):
        with column:
            st.markdown(_dash_card_markup(label, value, note=note, compact=compact), unsafe_allow_html=True)


def _formatar_percentual(value: float | None) -> str:
    if value is None:
        return "Indisponivel"
    return f"{value:.1f} %".replace(".", ",")


def _render_capacidade_banner(capacidade: CapacidadeOperacional) -> None:
    faixa = capacidade.faixa
    cor = faixa.cor_hex
    st.markdown(
        (
            f'<div class="gd-cap-banner" style="border-color:{cor};">'
            f'<span class="gd-cap-banner-title">Capacidade Operacional</span>'
            f'<span class="gd-cap-banner-status" style="color:{cor};">{html.escape(faixa.rotulo_banner)}</span>'
            f"</div>"
        ),
        unsafe_allow_html=True,
    )


def _render_capacidade_indicador(capacidade: CapacidadeOperacional) -> None:
    pct = capacidade.percentual
    pct_texto = f"{pct:.0f}%" if pct is not None else "—"
    cor = capacidade.faixa.cor_hex
    st.markdown(
        (
            f'<div class="gd-cap-bar-wrap">'
            f'<span>Capacidade</span>'
            f'<span class="gd-cap-bar" style="color:{cor};">{html.escape(capacidade.barra_visual)}</span>'
            f'<span style="color:{cor};">{html.escape(pct_texto)}</span>'
            f"</div>"
        ),
        unsafe_allow_html=True,
    )


def _render_alertas_capacidade(capacidade: CapacidadeOperacional) -> None:
    if capacidade.exibir_alerta_vermelho:
        st.error(
            "Capacidade operacional em risco de interrupcao (95% ou mais). "
            "Recomendamos executar a retencao do dia mais antigo. "
            "As importacoes continuam disponiveis — a decisao e do operador."
        )
    elif capacidade.exibir_aviso_discreto:
        st.warning(
            "Capacidade operacional em atencao (80% a 89%). "
            "Monitore o espaco e considere a retencao preventiva."
        )


def _render_banco_dados(uso: UsoBancoDados, capacidade: CapacidadeOperacional | None = None) -> None:
    utilizacao = _formatar_percentual(uso.utilizacao_percentual)
    _render_section("Banco de Dados")
    if capacidade is not None:
        _render_capacidade_indicador(capacidade)
    _render_cards_row(
        [
            ("Motor", str(uso.motor), ""),
            ("Espaco Utilizado", _formatar_bytes(uso.bytes_ocupados), ""),
            ("Espaco Livre", _formatar_bytes(uso.bytes_disponiveis), ""),
            ("Limite", _formatar_bytes(uso.bytes_limite), ""),
            ("Utilizacao", utilizacao, ""),
        ]
        )
    if uso.observacao:
        st.markdown(
            f'<p class="gd-dash-footnote">{html.escape(uso.observacao)}</p>',
            unsafe_allow_html=True,
        )


def _render_confirmacao_capacidade(
    previa: PreviaRetencaoCapacidade,
    confirmacao: ConfirmacaoRetencao,
) -> None:
    with st.container(border=True):
        _render_section("Retencao sugerida por capacidade")
        st.caption(
            f"Serao removidos os carregamentos do dia {_formatar_data_br(previa.data_alvo)}. "
            "A operacao e irreversivel."
        )
        _render_cards_row(
            [
                ("Data", _formatar_data_br(previa.data_alvo), ""),
                ("Carregamentos", f"{previa.carregamentos:,}", ""),
                ("Notas Fiscais", f"{previa.notas_fiscais:,}", ""),
                ("PDFs", f"{previa.documentos_pdf:,}", ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("XMLs", f"{previa.documentos_xml:,}", ""),
                ("Eventos", f"{previa.eventos:,}", ""),
                ("Historicos", f"{previa.historicos:,}", ""),
                ("Espaco recuperado (est.)", _formatar_bytes(previa.espaco_recuperavel_bytes), ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("Espaco atual", _formatar_bytes(previa.espaco_atual_bytes), ""),
                ("Espaco apos retencao", _formatar_bytes(previa.espaco_apos_bytes), ""),
                ("% atual", _formatar_percentual(previa.percentual_atual), ""),
                ("% apos retencao", _formatar_percentual(previa.percentual_apos), ""),
            ],
            compact=True,
        )
        st.error("Deseja confirmar a retencao deste dia?")

        col_confirmar, col_cancelar = st.columns(2)
        with col_confirmar:
            if st.button("Confirmar Retencao", type="primary", key="gestao_dados_capacidade_confirmar"):
                if not is_admin():
                    st.error("A retencao manual e restrita a administradores.")
                    return
                usuario = get_current_user()
                if usuario is None:
                    st.error("Usuario nao autenticado.")
                    return
                with st.spinner("Executando retencao transacional..."):
                    resultado = get_execucao_retencao_service().executar_retencao(
                        confirmacao,
                        usuario_id=int(usuario.id),
                        usuario_nome=str(usuario.nome or usuario.usuario),
                    )
                st.session_state[_SESSION_RESULTADO] = resultado
                st.session_state.pop(_SESSION_CAPACIDADE_PREVIA, None)
                st.session_state.pop(_SESSION_CAPACIDADE_CONFIRMACAO, None)
                st.session_state.pop(_SESSION_SIMULACAO, None)
                st.session_state.pop(_SESSION_CONFIRMACAO, None)
                if resultado.sucesso:
                    painel = get_gestao_dados_service().obter_painel()
                    if (painel.capacidade.percentual or 0) >= CAPACITY_ORANGE_MIN_PERCENT:
                        st.session_state[_SESSION_CAPACIDADE_INICIAR] = True
                st.rerun()
        with col_cancelar:
            if st.button("Cancelar", key="gestao_dados_capacidade_cancelar"):
                st.session_state.pop(_SESSION_CAPACIDADE_PREVIA, None)
                st.session_state.pop(_SESSION_CAPACIDADE_CONFIRMACAO, None)
                st.rerun()


def _iniciar_fluxo_capacidade_se_pendente() -> None:
    if not st.session_state.pop(_SESSION_CAPACIDADE_INICIAR, False):
        return
    try:
        with st.spinner("Montando previa do dia mais antigo..."):
            previa, confirmacao = get_gestao_capacidade_service().preparar_retencao_dia_mais_antigo()
    except (GestaoCapacidadeError, RetencaoExecucaoError) as exc:
        st.warning(str(exc))
        return
    st.session_state[_SESSION_CAPACIDADE_PREVIA] = previa
    st.session_state[_SESSION_CAPACIDADE_CONFIRMACAO] = confirmacao


def _render_controles_capacidade(capacidade: CapacidadeOperacional) -> None:
    if (capacidade.percentual or 0) < CAPACITY_ORANGE_MIN_PERCENT:
        return
    with st.container(border=True):
        _render_section("Retencao preventiva")
        st.caption(
            "Sugestao: excluir apenas o dia elegivel mais antigo, um dia por vez, "
            "para liberar espaco sem comprometer o periodo protegido."
        )
        if st.button("Executar Retencao do Dia Mais Antigo", key="gestao_dados_capacidade_iniciar"):
            if not is_admin():
                st.error("A retencao manual e restrita a administradores.")
                return
            st.session_state[_SESSION_CAPACIDADE_INICIAR] = True
            st.rerun()


def _render_politica_retencao(painel: GestaoDadosPainel) -> None:
    periodo = f"{_formatar_data_br(painel.periodo_inicio)} ate {_formatar_data_br(painel.periodo_fim)}"
    _render_section("Politica de Retencao")
    _render_cards_row(
        [
            ("Politica", str(painel.politica_descricao), ""),
            ("Periodo Mantido", f"{painel.politica_dias_mantidos} dias", ""),
            ("Status", str(painel.politica_status), ""),
            ("Periodo", periodo, ""),
        ]
    )


def _render_previa_retencao(painel: GestaoDadosPainel) -> None:
    pacote = painel.pacote
    _render_section("Previa da Retencao")
    _render_cards_row(
        [
            ("Carregamentos elegiveis", f"{pacote.carregamentos:,}", ""),
            ("Espaco recuperavel", _formatar_bytes(pacote.espaco_recuperavel_bytes), ""),
            ("PDFs", f"{pacote.documentos_pdf:,}", ""),
            ("XMLs", f"{pacote.documentos_xml:,}", ""),
            ("Eventos", f"{pacote.eventos:,}", ""),
            ("Historicos", f"{pacote.historicos:,}", ""),
        ],
        compact=True,
    )


def _render_contagens_elegiveis(painel: GestaoDadosPainel) -> None:
    pacote = painel.pacote
    _render_section("Elegiveis para retencao")
    _render_cards_row(
        [
            ("Notas Fiscais", f"{pacote.notas_fiscais:,}", ""),
            ("Itens", f"{pacote.itens_carregamento:,}", ""),
            ("Carregamentos", f"{pacote.carregamentos:,}", ""),
            ("XMLs", f"{pacote.documentos_xml:,}", ""),
        ],
        compact=True,
    )
    _render_cards_row(
        [
            ("PDFs", f"{pacote.documentos_pdf:,}", ""),
            ("Historicos", f"{pacote.historicos:,}", ""),
            ("Eventos", f"{pacote.eventos:,}", ""),
            ("Espaco recuperavel (est.)", _formatar_bytes(pacote.espaco_recuperavel_bytes), ""),
        ],
        compact=True,
    )
    st.markdown(
        f'<p class="gd-dash-footnote">Carregamentos com data anterior a '
        f"{_formatar_data_br(painel.data_corte)} sao classificados como elegiveis para retencao.</p>",
        unsafe_allow_html=True,
    )


def _render_simulacao_controles(painel: GestaoDadosPainel) -> None:
    with st.container(border=True):
        _render_section("Simulacao executavel")
        st.caption(
            "Percorre o fluxo de retencao validando pacotes, arquivos e dependencias. "
            "Nenhuma exclusao sera realizada nesta etapa."
        )
        if st.button("Executar Simulacao", type="primary", key="gestao_dados_executar_simulacao"):
            with st.spinner("Validando pacotes elegiveis..."):
                relatorio = get_simulacao_retencao_service().executar_simulacao()
            st.session_state[_SESSION_SIMULACAO] = relatorio
            st.session_state.pop(_SESSION_CONFIRMACAO, None)
            st.session_state.pop(_SESSION_RESULTADO, None)
            st.rerun()

        if not painel.possui_elegiveis:
            st.warning("Nao ha carregamentos elegiveis para simular neste momento.")


def _render_retencao_controles(relatorio: RelatorioSimulacaoRetencao) -> None:
    with st.container(border=True):
        _render_section("Retencao operacional")
        if not is_admin():
            st.caption("A retencao manual e restrita a administradores.")
            return

        aptos = relatorio.pacotes_apto_futura_retencao
        if not aptos:
            st.warning("Nenhum pacote apto para retencao. Corrija inconsistencias antes de continuar.")
            return

        if st.button("Executar Retencao", type="primary", key="gestao_dados_executar_retencao"):
            try:
                with st.spinner("Revalidando pacotes e montando confirmacao..."):
                    _, confirmacao = get_execucao_retencao_service().preparar_execucao()
            except RetencaoExecucaoError as exc:
                st.error(str(exc))
                return
            st.session_state[_SESSION_CONFIRMACAO] = confirmacao
            st.session_state.pop(_SESSION_RESULTADO, None)
            st.rerun()


def _render_confirmacao_retencao(confirmacao: ConfirmacaoRetencao) -> None:
    with st.container(border=True):
        _render_section("Retencao dos dados")
        st.caption("Serao removidos os registros abaixo. A operacao e irreversivel.")
        _render_cards_row(
            [
                ("Carregamentos", f"{confirmacao.carregamentos:,}", ""),
                ("Notas", f"{confirmacao.notas_fiscais:,}", ""),
                ("XML", f"{confirmacao.documentos_xml:,}", ""),
                ("PDFs", f"{confirmacao.documentos_pdf:,}", ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("Eventos", f"{confirmacao.eventos:,}", ""),
                ("Historicos", f"{confirmacao.historicos:,}", ""),
                ("Espaco estimado", _formatar_bytes(confirmacao.espaco_estimado_bytes), ""),
            ],
            compact=True,
        )
        st.error("Deseja continuar com a retencao?")

        col_confirmar, col_cancelar = st.columns(2)
        with col_confirmar:
            if st.button("Confirmar Retencao", type="primary", key="gestao_dados_confirmar_retencao"):
                usuario = get_current_user()
                if usuario is None:
                    st.error("Usuario nao autenticado.")
                    return
                with st.spinner("Executando retencao transacional..."):
                    resultado = get_execucao_retencao_service().executar_retencao(
                        confirmacao,
                        usuario_id=int(usuario.id),
                        usuario_nome=str(usuario.nome or usuario.usuario),
                    )
                st.session_state[_SESSION_RESULTADO] = resultado
                st.session_state.pop(_SESSION_CONFIRMACAO, None)
                st.session_state.pop(_SESSION_SIMULACAO, None)
                st.rerun()
        with col_cancelar:
            if st.button("Cancelar", key="gestao_dados_cancelar_retencao"):
                st.session_state.pop(_SESSION_CONFIRMACAO, None)
                st.rerun()


def _render_resultado_retencao(resultado: ResultadoRetencao) -> None:
    with st.container(border=True):
        _render_section("Resultado da retencao")
        if resultado.sucesso:
            st.success(resultado.mensagem)
            _render_cards_row(
                [
                    ("Carregamentos removidos", f"{resultado.carregamentos_removidos:,}", ""),
                    ("Notas removidas", f"{resultado.notas_fiscais_removidas:,}", ""),
                    ("XML removidos", f"{resultado.documentos_xml_removidos:,}", ""),
                    ("PDFs removidos", f"{resultado.documentos_pdf_removidos:,}", ""),
                ],
                compact=True,
            )
            _render_cards_row(
                [
                    ("Eventos removidos", f"{resultado.eventos_removidos:,}", ""),
                    ("Historicos removidos", f"{resultado.historicos_removidos:,}", ""),
                    ("Espaco recuperado", _formatar_bytes(resultado.espaco_recuperado_bytes), ""),
                    ("Tempo", f"{resultado.duracao_ms / 1000:.1f} s", ""),
                ],
                compact=True,
            )
            st.caption(
                f"Arquivos fisicos: {resultado.arquivos_pdf_removidos} PDF(s), "
                f"{resultado.arquivos_xml_removidos} XML(s) removidos apos commit."
            )
            if resultado.arquivos_falha:
                st.warning("Alguns arquivos nao puderam ser removidos do disco.")
        else:
            st.error(resultado.mensagem)
            if resultado.revertido:
                st.info("Nenhum dado foi removido. A transacao foi revertida (ROLLBACK).")


def _render_resumo_saude(relatorio: RelatorioSimulacaoRetencao) -> None:
    resumo = relatorio.resumo
    pdf_status = "Todos encontrados" if resumo.todos_pdfs_encontrados else "Arquivos ausentes detectados"
    xml_status = "Todos encontrados" if resumo.todos_xmls_encontrados else "Arquivos ausentes detectados"
    with st.container(border=True):
        _render_section("Saude da arvore")
        _render_cards_row(
            [
                ("Pacotes elegiveis", f"{resumo.pacotes_elegiveis:,}", ""),
                ("Pacotes integros", f"{resumo.pacotes_integros:,}", ""),
                ("Com inconsistencia", f"{resumo.pacotes_com_inconsistencia:,}", ""),
                ("Integridade geral", f"{resumo.integridade_geral_percentual:.1f} %", ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("Saudaveis", f"{resumo.pacotes_saudaveis:,}", ""),
                ("Atencao", f"{resumo.pacotes_atencao:,}", ""),
                ("Criticos", f"{resumo.pacotes_criticos:,}", ""),
            ],
            compact=True,
        )
        st.markdown(
            f'<p class="gd-dash-footnote"><strong>Arquivos PDF:</strong> {html.escape(pdf_status)} &nbsp;|&nbsp; '
            f"<strong>Arquivos XML:</strong> {html.escape(xml_status)}</p>",
            unsafe_allow_html=True,
        )


def _render_relatorio_final(relatorio: RelatorioSimulacaoRetencao) -> None:
    resumo = relatorio.resumo
    with st.container(border=True):
        _render_section("Simulacao concluida")
        st.info("Simulacao concluida sem alteracoes no banco.")
        _render_cards_row(
            [
                ("Pacotes analisados", f"{resumo.pacotes_elegiveis:,}", ""),
                ("Pacotes integros", f"{resumo.pacotes_integros:,}", ""),
                ("Com inconsistencia", f"{resumo.pacotes_com_inconsistencia:,}", ""),
                ("Registros analisados", f"{relatorio.registros_analisados:,}", ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("Arquivos PDF", f"{relatorio.arquivos_pdf:,}", ""),
                ("Arquivos XML", f"{relatorio.arquivos_xml:,}", ""),
                ("Espaco elegivel", _formatar_bytes(relatorio.espaco_elegivel_bytes), ""),
                ("Tempo", f"{relatorio.duracao_ms / 1000:.1f} s", ""),
            ],
            compact=True,
        )
        aptos = len(relatorio.pacotes_apto_futura_retencao)
        if aptos == resumo.pacotes_elegiveis:
            st.caption("Todos os pacotes elegiveis estao aptos para retencao manual.")
        else:
            st.warning(
                f"{resumo.pacotes_elegiveis - aptos} pacote(s) deverao ser corrigidos antes da retencao."
            )


def _render_auditoria_tecnica(relatorio: RelatorioSimulacaoRetencao) -> None:
    with st.expander("Relatorio tecnico de auditoria", expanded=False):
        if relatorio.orfaos:
            st.markdown("**Orfaos e vinculos globais**")
            for problema in relatorio.orfaos:
                st.markdown(
                    f"- [{problema.severidade.emoji} {problema.severidade.rotulo}] "
                    f"{html.escape(problema.descricao)}"
                )

        inconsistentes = [p for p in relatorio.pacotes if p.saude != SaudePacote.SAUDAVEL]
        if inconsistentes:
            st.markdown("**Carregamentos com inconsistencias**")
            for pacote in inconsistentes:
                st.markdown(
                    f"- {pacote.numero_carregamento} ({_formatar_data_br(pacote.data_carregamento)}): "
                    f"{pacote.saude.emoji} {pacote.saude.rotulo}"
                )
                for problema in pacote.problemas:
                    st.markdown(f"  - {html.escape(problema)}")

        aptos = relatorio.pacotes_apto_futura_retencao
        if aptos:
            st.markdown("**Pacotes aptos para retencao**")
            numeros = ", ".join(p.numero_carregamento for p in aptos)
            st.caption(numeros)

        corrigir = relatorio.pacotes_requerem_correcao
        if corrigir:
            st.markdown("**Pacotes que exigem correcao antes da retencao**")
            for pacote in corrigir:
                st.markdown(
                    f"- {pacote.numero_carregamento}: {pacote.saude.emoji} "
                    f"{'; '.join(pacote.problemas) or 'Dependencia critica'}"
                )


def _render_detalhe_pacote(pacote: PacoteRetencaoUnitario) -> None:
    with st.container(border=True):
        _render_section("Detalhe do carregamento")
        st.caption(
            f"Carregamento {pacote.numero_carregamento} | "
            f"Data {_formatar_data_br(pacote.data_carregamento)} | "
            f"Saude {pacote.saude.emoji} {pacote.saude.rotulo}"
        )
        _render_cards_row(
            [
                ("Itens", str(pacote.itens_carregamento), ""),
                ("Notas", str(pacote.notas_fiscais), ""),
                ("Documento XML", str(pacote.documentos_xml), ""),
                ("PDF", str(pacote.documentos_pdf), ""),
            ],
            compact=True,
        )
        _render_cards_row(
            [
                ("Eventos", str(pacote.eventos), ""),
                ("Historicos", str(pacote.historicos), ""),
                ("Arquivos encontrados", str(pacote.arquivos_encontrados), ""),
                ("Arquivos ausentes", str(pacote.arquivos_ausentes), ""),
            ],
            compact=True,
        )
        _render_cards_row([("Integridade", f"{pacote.integridade_percentual:.1f} %", "")], compact=True)
        if pacote.apto_retencao:
            st.success("Este carregamento esta apto para retencao.")
        else:
            st.error("Este carregamento nao esta apto para retencao por dependencia critica.")
        if pacote.problemas:
            st.markdown("**Observacoes**")
            for problema in pacote.problemas:
                st.markdown(f"- {html.escape(problema)}")


def _render_seletor_pacotes(relatorio: RelatorioSimulacaoRetencao) -> None:
    if not relatorio.pacotes:
        return
    opcoes = {
        f"{p.numero_carregamento} — {_formatar_data_br(p.data_carregamento)} ({p.saude.emoji})": p.carregamento_id
        for p in relatorio.pacotes
    }
    labels = list(opcoes.keys())
    selecionado = st.selectbox(
        "Selecionar carregamento para detalhe",
        options=labels,
        key="gestao_dados_select_pacote",
    )
    pacote_id = opcoes[selecionado]
    pacote = next(p for p in relatorio.pacotes if p.carregamento_id == pacote_id)
    _render_detalhe_pacote(pacote)


def render_gestao_dados_page(*, render_header_callback) -> None:
    render_header_callback(
        "Gestao de Dados",
        "Dashboard administrativo do banco e da politica de retencao operacional.",
    )

    _inject_gestao_dados_styles()
    st.markdown('<div class="gd-page-root">', unsafe_allow_html=True)

    _iniciar_fluxo_capacidade_se_pendente()

    painel = get_gestao_dados_service().obter_painel()
    relatorio: RelatorioSimulacaoRetencao | None = st.session_state.get(_SESSION_SIMULACAO)
    confirmacao: ConfirmacaoRetencao | None = st.session_state.get(_SESSION_CONFIRMACAO)
    confirmacao_capacidade: ConfirmacaoRetencao | None = st.session_state.get(_SESSION_CAPACIDADE_CONFIRMACAO)
    previa_capacidade: PreviaRetencaoCapacidade | None = st.session_state.get(_SESSION_CAPACIDADE_PREVIA)
    resultado: ResultadoRetencao | None = st.session_state.get(_SESSION_RESULTADO)

    _render_capacidade_banner(painel.capacidade)
    _render_alertas_capacidade(painel.capacidade)

    if resultado is not None:
        _render_resultado_retencao(resultado)

    _render_banco_dados(painel.uso_banco, painel.capacidade)
    _render_politica_retencao(painel)
    _render_controles_capacidade(painel.capacidade)

    if confirmacao_capacidade is not None and previa_capacidade is not None and is_admin():
        _render_confirmacao_capacidade(previa_capacidade, confirmacao_capacidade)
        st.markdown("</div>", unsafe_allow_html=True)
        return

    if confirmacao is not None and is_admin():
        _render_confirmacao_retencao(confirmacao)
        st.markdown("</div>", unsafe_allow_html=True)
        return

    if painel.possui_elegiveis:
        _render_previa_retencao(painel)
        _render_contagens_elegiveis(painel)
        _render_simulacao_controles(painel)

        if relatorio is not None:
            _render_resumo_saude(relatorio)
            _render_relatorio_final(relatorio)
            _render_retencao_controles(relatorio)
            _render_seletor_pacotes(relatorio)
            _render_auditoria_tecnica(relatorio)
    else:
        st.markdown('<div style="margin-top: 24px;" aria-hidden="true"></div>', unsafe_allow_html=True)
        st.success(
            f"Nenhum carregamento elegivel para retencao. "
            f"Periodo mantido: {painel.politica_descricao} "
            f"({_formatar_data_br(painel.periodo_inicio)} a {_formatar_data_br(painel.periodo_fim)})."
        )

    if resultado is None and relatorio is None:
        st.caption("Execute a simulacao para validar os pacotes elegiveis antes da retencao.")

    st.markdown("</div>", unsafe_allow_html=True)
