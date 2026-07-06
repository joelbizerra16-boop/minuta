from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from carregamentos.integration import (
    OPERACIONAL_ANALISE_CONFIRMADA_KEY,
    OPERACIONAL_CONTINUACAO_AUDITORIA_KEY,
    OPERACIONAL_CONTINUAR_HISTORICO_VALUE,
    OPERACIONAL_DECISAO_WIDGET_KEY,
    cancelar_operacao_pendente,
    confirmar_analise_operacional_continuacao,
    confirmar_decisao_operacional_continuacao,
    get_diagnostico_efetivo_fechamento,
    get_operacional_decisao,
    inferir_decisao_operacional,
    requer_confirmacao_explicita_historico,
    resolve_operational_panel_mode,
    set_operacional_diagnostico,
)
from carregamentos.models.operacional import (
    CenarioOperacional,
    DecisaoOperacional,
    DiagnosticoCarregamento,
)


class _SessionState(dict):
    def get(self, key, default=None):
        return super().get(key, default)

    def pop(self, key, default=None):
        return super().pop(key, default)


@pytest.fixture
def mock_streamlit_session(monkeypatch):
    state = _SessionState()
    mock_st = MagicMock()
    mock_st.session_state = state
    monkeypatch.setattr("carregamentos.integration.st", mock_st)
    return state


def _diagnostico_reimpressao() -> DiagnosticoCarregamento:
    return DiagnosticoCarregamento(
        cenario=CenarioOperacional.REIMPRESSAO,
        nfs_total=18,
        nfs_novas=0,
        nfs_existentes=18,
        carregamentos_distintos=1,
        requer_decisao=True,
        opcoes_decisao=[
            DecisaoOperacional.REIMPRIMIR,
            DecisaoOperacional.REENTREGA,
            DecisaoOperacional.CANCELAR,
        ],
    )


def _diagnostico_conflito() -> DiagnosticoCarregamento:
    return DiagnosticoCarregamento(
        cenario=CenarioOperacional.CONFLITO_MULTIPLO,
        nfs_total=18,
        nfs_novas=0,
        nfs_existentes=18,
        carregamentos_distintos=2,
        requer_decisao=True,
        bloqueia_fechamento=True,
        opcoes_decisao=[DecisaoOperacional.CANCELAR],
    )


def test_resolve_mode_historico_antes_da_confirmacao(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "carregamento_historico"


def test_resolve_mode_decisao_apos_confirmacao(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    confirmar_analise_operacional_continuacao()
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "carregamento_decisao"


def test_confirmar_analise_limpa_decisao_parcial(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.CANCELAR.value
    confirmar_analise_operacional_continuacao()
    assert mock_streamlit_session[OPERACIONAL_ANALISE_CONFIRMADA_KEY] is True
    assert OPERACIONAL_DECISAO_WIDGET_KEY not in mock_streamlit_session
    assert "operacional_decisao" not in mock_streamlit_session


def test_resolve_mode_fechamento_apos_selecao_no_widget(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    confirmar_analise_operacional_continuacao()
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.REIMPRIMIR.value
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "fechamento"
    assert mock_streamlit_session["operacional_decisao"] == DecisaoOperacional.REIMPRIMIR.value


def test_requer_confirmacao_explicita_para_conflito(mock_streamlit_session) -> None:
    diagnostico = _diagnostico_conflito()
    assert requer_confirmacao_explicita_historico(diagnostico) is True


def test_confirmacao_explicita_registra_auditoria_e_libera_fechamento(mock_streamlit_session) -> None:
    diagnostico = _diagnostico_conflito()
    set_operacional_diagnostico(diagnostico)
    confirmar_analise_operacional_continuacao()
    mock_streamlit_session["summary"] = {"numero_carga": "000099"}
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = OPERACIONAL_CONTINUAR_HISTORICO_VALUE
    decisao = confirmar_decisao_operacional_continuacao()
    assert decisao == DecisaoOperacional.NOVO
    registro = mock_streamlit_session[OPERACIONAL_CONTINUACAO_AUDITORIA_KEY]
    assert registro["carregamento_atual"] == "000099"
    assert registro["quantidade_nfs_historico"] == 18
    assert "autorizou manualmente" in registro["decisao"]
    efetivo = get_diagnostico_efetivo_fechamento()
    assert efetivo is not None
    assert efetivo.bloqueia_fechamento is False
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "fechamento"


def test_inferir_decisao_conflito_usa_novo(mock_streamlit_session) -> None:
    assert inferir_decisao_operacional(_diagnostico_conflito()) == DecisaoOperacional.NOVO


def test_get_operacional_decisao_sincroniza_widget(mock_streamlit_session) -> None:
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.REIMPRIMIR.value
    assert get_operacional_decisao() == DecisaoOperacional.REIMPRIMIR
    assert mock_streamlit_session["operacional_decisao"] == DecisaoOperacional.REIMPRIMIR.value


def test_cancelar_operacao_limpa_confirmacao_analise(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    confirmar_analise_operacional_continuacao()
    cancelar_operacao_pendente()
    assert OPERACIONAL_ANALISE_CONFIRMADA_KEY not in mock_streamlit_session
    assert OPERACIONAL_CONTINUACAO_AUDITORIA_KEY not in mock_streamlit_session
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "carregamento_historico"
