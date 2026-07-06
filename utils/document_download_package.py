from __future__ import annotations

import re
import zipfile
from io import BytesIO


def sanitize_filename_part(value: object, default: str) -> str:
    text = str(value or "").strip() or default
    text = re.sub(r"[^\w.\-]+", "_", text, flags=re.UNICODE)
    return text.strip("._") or default


XML_MISSING_WARNING = "Alguns XMLs nao estao disponiveis para exportacao."


def build_documentos_download_package(
    *,
    carregamento_pdf_bytes: bytes | None,
    entrega_pdf_bytes: bytes | None,
    carregamento_selected: bool,
    entrega_selected: bool,
    xml_selected: bool,
    xml_entries: list[tuple[str, bytes]] | None,
    numero_carga: str,
) -> tuple[bytes, str, str, str]:
    pdf_minuta = carregamento_pdf_bytes if carregamento_selected else None
    pdf_romaneio = entrega_pdf_bytes if entrega_selected else None
    xml_files = list(xml_entries or []) if xml_selected else []

    has_minuta = bool(pdf_minuta)
    has_romaneio = bool(pdf_romaneio)
    has_xml = bool(xml_files)

    if not carregamento_selected and not entrega_selected and not xml_selected:
        return b"", "", "application/pdf", "Selecione ao menos um tipo de documento para gerar o download"

    warnings: list[str] = []
    if xml_selected and not has_xml:
        warnings.append(XML_MISSING_WARNING)

    carga_slug = sanitize_filename_part(numero_carga, "brida")
    includes_pdf = carregamento_selected or entrega_selected
    includes_xml = xml_selected

    if includes_xml and includes_pdf:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            if carregamento_selected:
                if has_minuta:
                    archive.writestr("Minuta.pdf", pdf_minuta)
                else:
                    warnings.append("Nao ha dados disponiveis para gerar a minuta de carregamento.")
            if entrega_selected:
                if has_romaneio:
                    archive.writestr("Romaneio.pdf", pdf_romaneio)
                else:
                    warnings.append("Nao ha dados validos disponiveis para gerar o romaneio de entrega.")
            for filename, content in xml_files:
                archive.writestr(f"XML/{filename}", content)
        payload = zip_buffer.getvalue()
        if not _zip_has_files(payload):
            return b"", "", "application/pdf", _join_warnings(warnings) or (
                "Nao foi possivel gerar os documentos selecionados."
            )
        return payload, f"Carregamento_{carga_slug}.zip", "application/zip", _join_warnings(warnings)

    if includes_xml and not includes_pdf:
        if not has_xml:
            return b"", "", "application/pdf", _join_warnings(warnings) or XML_MISSING_WARNING
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            for filename, content in xml_files:
                archive.writestr(f"XML/{filename}", content)
        return (
            zip_buffer.getvalue(),
            f"Carregamento_{carga_slug}.zip",
            "application/zip",
            _join_warnings(warnings),
        )

    if carregamento_selected and not entrega_selected:
        if not pdf_minuta:
            return b"", "", "application/pdf", "Nao ha dados disponiveis para gerar a minuta de carregamento."
        return (
            pdf_minuta,
            f"minuta_carregamento_{carga_slug}.pdf",
            "application/pdf",
            _join_warnings(warnings),
        )

    if entrega_selected and not carregamento_selected:
        if not pdf_romaneio:
            return b"", "", "application/pdf", "Nao ha dados validos disponiveis para gerar o romaneio de entrega."
        return (
            pdf_romaneio,
            f"minuta_entrega_{carga_slug}.pdf",
            "application/pdf",
            _join_warnings(warnings),
        )

    if has_minuta and has_romaneio:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("minuta_carregamento.pdf", pdf_minuta)
            archive.writestr("minuta_entrega.pdf", pdf_romaneio)
        return (
            zip_buffer.getvalue(),
            f"minutas_{carga_slug}.zip",
            "application/zip",
            _join_warnings(warnings),
        )
    if has_minuta:
        return (
            pdf_minuta,
            f"minuta_carregamento_{carga_slug}.pdf",
            "application/pdf",
            "Nao foi possivel gerar o romaneio de entrega. A minuta de carregamento esta disponivel para download.",
        )
    if has_romaneio:
        return (
            pdf_romaneio,
            f"minuta_entrega_{carga_slug}.pdf",
            "application/pdf",
            "Nao foi possivel gerar a minuta de carregamento. O romaneio de entrega esta disponivel para download.",
        )
    return (
        b"",
        "",
        "application/pdf",
        _join_warnings(warnings) or "Nao foi possivel gerar a minuta de carregamento e o romaneio de entrega.",
    )


def _zip_has_files(payload: bytes) -> bool:
    if not payload:
        return False
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        return bool(archive.namelist())


def _join_warnings(warnings: list[str]) -> str:
    unique: list[str] = []
    for item in warnings:
        text = str(item or "").strip()
        if text and text not in unique:
            unique.append(text)
    return ". ".join(unique)
