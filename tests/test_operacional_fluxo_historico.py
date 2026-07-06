from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from carregamentos.integration import (
    OPERACIONAL_ANALISE_CONFIRMADA_KEY,
    OPERACIONAL_DECISAO_WIDGET_KEY,
    cancelar_operacao_pendente,
    confirmar_analise_operacional_continuacao,
    get_operacional_decisao,
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
        requer_decisao=True,
        opcoes_decisao=[
            DecisaoOperacional.REIMPRIMIR,
            DecisaoOperacional.REENTREGA,
            DecisaoOperacional.CANCELAR,
        ],
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


def test_get_operacional_decisao_sincroniza_widget(mock_streamlit_session) -> None:
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.REIMPRIMIR.value
    assert get_operacional_decisao() == DecisaoOperacional.REIMPRIMIR
    assert mock_streamlit_session["operacional_decisao"] == DecisaoOperacional.REIMPRIMIR.value


def test_cancelar_operacao_limpa_confirmacao_analise(mock_streamlit_session) -> None:
    set_operacional_diagnostico(_diagnostico_reimpressao())
    confirmar_analise_operacional_continuacao()
    cancelar_operacao_pendente()
    assert OPERACIONAL_ANALISE_CONFIRMADA_KEY not in mock_streamlit_session
    mode = resolve_operational_panel_mode(has_excel_loaded=True, processed_df_empty=False)
    assert mode == "carregamento_historico"
