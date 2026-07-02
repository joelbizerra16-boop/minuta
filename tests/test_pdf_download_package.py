from __future__ import annotations

import zipfile
from io import BytesIO

from app import build_pdf_download_package


def _call_both(minuta: bytes | None, romaneio: bytes | None) -> tuple[bytes, str, str, str]:
    return build_pdf_download_package(
        carregamento_pdf_bytes=minuta,
        entrega_pdf_bytes=romaneio,
        carregamento_selected=True,
        entrega_selected=True,
        numero_carga="000005",
    )


def test_cenario_1_ambos_ok() -> None:
    payload, name, mime, message = _call_both(b"%PDF-minuta%", b"%PDF-romaneio%")
    assert payload
    assert name.endswith(".zip")
    assert mime == "application/zip"
    assert message == ""
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert "minuta_carregamento.pdf" in archive.namelist()
        assert "minuta_entrega.pdf" in archive.namelist()
    print("cenario 1 OK")


def test_cenario_2_minuta_ok_romaneio_erro() -> None:
    payload, name, mime, message = _call_both(b"%PDF-minuta%", None)
    assert payload == b"%PDF-minuta%"
    assert name.startswith("minuta_carregamento_")
    assert mime == "application/pdf"
    assert "romaneio de entrega" in message.lower()
    assert "minuta de carregamento esta disponivel" in message.lower()
    print("cenario 2 OK")


def test_cenario_3_minuta_erro_romaneio_ok() -> None:
    payload, name, mime, message = _call_both(None, b"%PDF-romaneio%")
    assert payload == b"%PDF-romaneio%"
    assert name.startswith("minuta_entrega_")
    assert mime == "application/pdf"
    assert "minuta de carregamento" in message.lower()
    assert "romaneio de entrega esta disponivel" in message.lower()
    print("cenario 3 OK")


def test_cenario_4_ambos_falharam() -> None:
    payload, name, mime, message = _call_both(None, None)
    assert payload == b""
    assert name == ""
    assert "carregamento" in message.lower()
    assert "romaneio" in message.lower()
    print("cenario 4 OK")


def test_minuta_isolada() -> None:
    payload, name, mime, message = build_pdf_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=None,
        carregamento_selected=True,
        entrega_selected=False,
        numero_carga="000005",
    )
    assert payload == b"%PDF-minuta%"
    assert message == ""
    print("minuta isolada OK")


def test_romaneio_isolado() -> None:
    payload, name, mime, message = build_pdf_download_package(
        carregamento_pdf_bytes=None,
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=False,
        entrega_selected=True,
        numero_carga="000005",
    )
    assert payload == b"%PDF-romaneio%"
    assert message == ""
    print("romaneio isolado OK")


if __name__ == "__main__":
    test_cenario_1_ambos_ok()
    test_cenario_2_minuta_ok_romaneio_erro()
    test_cenario_3_minuta_erro_romaneio_ok()
    test_cenario_4_ambos_falharam()
    test_minuta_isolada()
    test_romaneio_isolado()
    print("All pdf download package tests passed")
