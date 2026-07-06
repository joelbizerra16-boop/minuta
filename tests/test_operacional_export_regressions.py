from __future__ import annotations

import zipfile
from io import BytesIO
from unittest.mock import MagicMock

import pytest

from carregamentos.integration import (
    OPERACIONAL_DECISAO_WIDGET_KEY,
    confirmar_decisao_operacional_continuacao,
    on_baixar_pdf_click,
    snapshot_exportacao_documentos,
)
from carregamentos.models.operacional import DecisaoOperacional
from utils.document_download_package import build_documentos_download_package


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


def test_confirmar_decisao_operacional_continuacao_normaliza_valor(mock_streamlit_session) -> None:
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.REIMPRIMIR.value
    decisao = confirmar_decisao_operacional_continuacao()
    assert decisao == DecisaoOperacional.REIMPRIMIR
    assert mock_streamlit_session["operacional_decisao"] == DecisaoOperacional.REIMPRIMIR.value


def test_confirmar_decisao_sem_widget_preserva_decisao_confirmada(mock_streamlit_session) -> None:
    mock_streamlit_session["operacional_decisao"] = DecisaoOperacional.COMPLEMENTAR.value
    decisao = confirmar_decisao_operacional_continuacao()
    assert decisao == DecisaoOperacional.COMPLEMENTAR


def test_snapshot_exportacao_documentos_respeita_checkboxes(mock_streamlit_session) -> None:
    mock_streamlit_session["minuta_pdf_carregamento"] = False
    mock_streamlit_session["minuta_pdf_entrega"] = True
    mock_streamlit_session["minuta_pdf_xmls"] = False
    snapshot = snapshot_exportacao_documentos()
    assert snapshot == {
        "carregamento_selected": False,
        "entrega_selected": True,
        "xml_selected": False,
    }


def test_on_baixar_pdf_click_nao_congela_exportacao_no_callback(mock_streamlit_session) -> None:
    mock_streamlit_session[OPERACIONAL_DECISAO_WIDGET_KEY] = DecisaoOperacional.REIMPRIMIR.value
    mock_streamlit_session["minuta_pdf_carregamento"] = False
    mock_streamlit_session["minuta_pdf_entrega"] = True
    mock_streamlit_session["minuta_pdf_xmls"] = True
    mock_streamlit_session["pdf_download_payload"] = b"stale"

    on_baixar_pdf_click()

    action = mock_streamlit_session["_processing_action"]
    assert action["type"] == "baixar_pdf"
    assert action["confirmar_reimpressao"] is True
    assert "carregamento_selected" not in action
    assert "entrega_selected" not in action
    assert "xml_selected" not in action
    assert "pdf_download_payload" not in mock_streamlit_session


def test_build_documentos_romaneio_mais_xml_sem_minuta() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=False,
        entrega_selected=True,
        xml_selected=True,
        xml_entries=[("nf.xml", b"<xml/>")],
        numero_carga="000019",
    )
    assert mime == "application/zip"
    assert name == "Carregamento_000019.zip"
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        names = archive.namelist()
        assert "Minuta.pdf" not in names
        assert "Romaneio.pdf" in names
        assert names == ["Romaneio.pdf", "XML/nf.xml"]
    assert message == ""


def test_build_documentos_todos_selecionados() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=True,
        entrega_selected=True,
        xml_selected=True,
        xml_entries=[("nf.xml", b"<xml/>")],
        numero_carga="000019",
    )
    assert mime == "application/zip"
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        names = set(archive.namelist())
        assert names == {"Minuta.pdf", "Romaneio.pdf", "XML/nf.xml"}
    assert message == ""


def test_build_documentos_apenas_minuta() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=True,
        entrega_selected=False,
        xml_selected=False,
        xml_entries=[("nf.xml", b"<xml/>")],
        numero_carga="000123",
    )
    assert payload == b"%PDF-minuta%"
    assert mime == "application/pdf"
    assert name.startswith("minuta_carregamento_")
    assert message == ""


def test_build_documentos_apenas_romaneio() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=False,
        entrega_selected=True,
        xml_selected=False,
        xml_entries=None,
        numero_carga="000123",
    )
    assert payload == b"%PDF-romaneio%"
    assert mime == "application/pdf"
    assert name.startswith("minuta_entrega_")
    assert message == ""


def test_build_documentos_apenas_xml() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=False,
        entrega_selected=False,
        xml_selected=True,
        xml_entries=[("nf.xml", b"<xml/>")],
        numero_carga="000123",
    )
    assert mime == "application/zip"
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert archive.namelist() == ["XML/nf.xml"]
    assert message == ""


def test_build_documentos_minuta_mais_romaneio_sem_xml() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=True,
        entrega_selected=True,
        xml_selected=False,
        xml_entries=None,
        numero_carga="000123",
    )
    assert mime == "application/zip"
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert set(archive.namelist()) == {"minuta_carregamento.pdf", "minuta_entrega.pdf"}
    assert message == ""
