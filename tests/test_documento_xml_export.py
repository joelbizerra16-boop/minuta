from __future__ import annotations

import hashlib
import os
import zipfile
from datetime import date, time
from decimal import Decimal
from io import BytesIO
from pathlib import Path

import pytest

from auth.bootstrap import configure_auth_storage
from carregamentos.services.xml_export_service import XmlExportService
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.schema import ensure_full_schema
from infrastructure.services.documento_xml_service import DocumentoXmlService, XmlDocumentalItem
from infrastructure.storage.xml_chave_extractor import extract_chave_nfe_from_xml_bytes
from utils.document_download_package import XML_MISSING_WARNING, build_documentos_download_package


CHAVE = "35260612345678901234550010000012345678901234"
CHAVE_2 = "35260612345678901234550010000098765432109876"
CHAVE_OUT = "35260612345678901234550010000011111111111111"


def _sample_xml(chave: str, numero: str = "1426799") -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc><NFe><infNFe><ide><nNF>{numero}</nNF></ide><chNFe>{chave}</chNFe></infNFe></NFe></nfeProc>""".encode(
        "utf-8"
    )


@pytest.fixture()
def xml_env(tmp_path: Path):
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(tmp_path)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_path / 'documento_xml.db').as_posix()}"
    get_settings.cache_clear()
    storage_dir = tmp_path / "xml_storage"
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=tmp_path,
        pdf_storage_dir=tmp_path / "documentos",
        xml_storage_dir=storage_dir,
    )
    ensure_full_schema()
    configure_auth_storage(tmp_path)
    yield storage_dir
    get_settings.cache_clear()


def test_extract_chave_from_bytes() -> None:
    payload = _sample_xml(CHAVE)
    assert extract_chave_nfe_from_xml_bytes(payload) == CHAVE


def test_persist_raw_xml_batch_single_commit(xml_env: Path) -> None:
    from sqlalchemy import event
    from sqlalchemy.orm import Session

    from infrastructure.database import get_engine

    service = DocumentoXmlService(storage_dir=xml_env)
    payload_a = _sample_xml(CHAVE, "1426799")
    payload_b = _sample_xml(CHAVE_2, "1426804")
    items = [
        XmlDocumentalItem(
            file_bytes=payload_a,
            hash_sha256=hashlib.sha256(payload_a).hexdigest(),
            original_filename="NF1426799.xml",
            chave_nfe=CHAVE,
        ),
        XmlDocumentalItem(
            file_bytes=payload_b,
            hash_sha256=hashlib.sha256(payload_b).hexdigest(),
            original_filename="NF1426804.xml",
            chave_nfe=CHAVE_2,
        ),
    ]
    stats = {"flush": 0, "commit": 0}

    @event.listens_for(Session, "after_flush")
    def _after_flush(session, flush_context) -> None:
        stats["flush"] += 1

    @event.listens_for(get_engine(), "commit")
    def _on_commit(conn) -> None:
        stats["commit"] += 1

    result = service.persist_raw_xml_batch(items, usuario_id=1)
    assert result.saved == 2
    assert len(list(xml_env.glob("*.xml"))) == 2
    assert stats["flush"] == 1
    assert stats["commit"] == 1


def test_persist_xml_dedup_by_hash(xml_env: Path) -> None:
    service = DocumentoXmlService(storage_dir=xml_env)
    payload = _sample_xml(CHAVE)
    first = service.persist_raw_xml(payload, original_filename="NF1426799.xml", usuario_id=1)
    second = service.persist_raw_xml(payload, original_filename="NF1426799.xml", usuario_id=1)
    assert first is not None and first.status == "saved"
    assert second is not None and second.status == "reused"
    assert len(list(xml_env.glob("*.xml"))) == 1


def test_xml_content_survives_local_file_loss(xml_env: Path) -> None:
    from infrastructure.repositories.sql.documento_xml_repository import SqlDocumentoXmlRepository
    from infrastructure.unit_of_work import UnitOfWork

    service = DocumentoXmlService(storage_dir=xml_env)
    payload = _sample_xml(CHAVE)
    result = service.persist_raw_xml(payload, original_filename="NF1426799.xml", usuario_id=1)
    assert result is not None and result.status == "saved"

    with UnitOfWork() as uow:
        record = SqlDocumentoXmlRepository(uow.session).get_by_chave(CHAVE)
    assert record is not None
    assert record.conteudo_xml == payload

    (xml_env / f"{CHAVE}.xml").unlink()
    assert service.read_xml_bytes(record) == payload


def test_database_persistence_does_not_depend_on_local_disk(xml_env: Path) -> None:
    from infrastructure.repositories.sql.documento_xml_repository import SqlDocumentoXmlRepository
    from infrastructure.unit_of_work import UnitOfWork

    blocked_storage = xml_env.parent / "storage_is_a_file"
    blocked_storage.write_bytes(b"not-a-directory")
    service = DocumentoXmlService(storage_dir=blocked_storage)
    payload = _sample_xml(CHAVE)

    result = service.persist_raw_xml(payload, original_filename="NF1426799.xml", usuario_id=1)
    assert result is not None and result.status == "saved"

    with UnitOfWork() as uow:
        record = SqlDocumentoXmlRepository(uow.session).get_by_chave(CHAVE)
    assert record is not None
    assert record.conteudo_xml == payload
    assert service.read_xml_bytes(record) == payload


def test_reading_legacy_disk_xml_backfills_database(xml_env: Path) -> None:
    from infrastructure.models.documento_xml import DocumentoXmlORM
    from infrastructure.repositories.sql.documento_xml_repository import SqlDocumentoXmlRepository
    from infrastructure.unit_of_work import UnitOfWork

    service = DocumentoXmlService(storage_dir=xml_env)
    payload = _sample_xml(CHAVE)
    service.persist_raw_xml(payload, original_filename="NF1426799.xml", usuario_id=1)

    with UnitOfWork() as uow:
        row = uow.session.query(DocumentoXmlORM).filter_by(chave_nfe=CHAVE).one()
        row.conteudo_xml = None

    with UnitOfWork() as uow:
        legacy_record = SqlDocumentoXmlRepository(uow.session).get_by_chave(CHAVE)
    assert legacy_record is not None and legacy_record.conteudo_xml is None
    assert service.read_xml_bytes(legacy_record) == payload

    (xml_env / f"{CHAVE}.xml").unlink()
    with UnitOfWork() as uow:
        durable_record = SqlDocumentoXmlRepository(uow.session).get_by_chave(CHAVE)
    assert durable_record is not None and durable_record.conteudo_xml == payload
    assert service.read_xml_bytes(durable_record) == payload


def test_persist_xml_single_physical_file_for_same_chave(xml_env: Path) -> None:
    service = DocumentoXmlService(storage_dir=xml_env)
    payload_a = _sample_xml(CHAVE, "1426799")
    payload_b = _sample_xml(CHAVE, "1426799").replace(b"1426799", b"9999999")
    service.persist_raw_xml(payload_a, original_filename="NF1426799.xml")
    service.persist_raw_xml(payload_b, original_filename="NF1426799.xml")
    files = list(xml_env.glob("*.xml"))
    assert len(files) == 1
    assert files[0].read_bytes() == payload_b


_seed_counter = 0


def _seed_carregamento(chaves: list[str], numero: str | None = None) -> int:
    global _seed_counter
    from infrastructure.database import get_session_factory

    _seed_counter += 1
    numero_carregamento = numero or f"0000{_seed_counter:03d}"

    session = get_session_factory()()
    try:
        carregamento = CarregamentoORM(
            numero_carregamento=numero_carregamento,
            usuario_id=1,
            data=date(2026, 7, 3),
            hora=time(10, 0, 0),
            motorista="Motorista",
            placa="ABC1D23",
            filial="BRIDA",
            data_saida="2026-07-03",
            modalidade="VEICULO",
            status="FINALIZADO",
            quantidade_nf=len(chaves),
            quantidade_itens=len(chaves),
            peso_total=Decimal("0"),
        )
        session.add(carregamento)
        session.flush()
        for index, chave in enumerate(chaves, start=1):
            session.add(
                ItemCarregamentoORM(
                    carregamento_id=int(carregamento.id),
                    numero_nf=chave[25:34].lstrip("0") or chave[25:34],
                    chave_nfe=chave,
                    codigo_produto="--",
                    sequencia=index,
                )
            )
        session.commit()
        return int(carregamento.id)
    finally:
        session.close()


def test_export_only_carregamento_xmls(xml_env: Path) -> None:
    service = DocumentoXmlService(storage_dir=xml_env)
    service.persist_raw_xml(_sample_xml(CHAVE), original_filename="NF1426799.xml")
    service.persist_raw_xml(_sample_xml(CHAVE_2), original_filename="NF1426804.xml")
    service.persist_raw_xml(_sample_xml(CHAVE_OUT), original_filename="NF999.xml")

    carregamento_id = _seed_carregamento([CHAVE, CHAVE_2])
    export = XmlExportService(documento_xml_service=service).collect_xmls_for_carregamento(carregamento_id)
    assert len(export.entries) == 2
    assert {entry.chave_nfe for entry in export.entries} == {CHAVE, CHAVE_2}


def test_export_missing_xml_does_not_fail(xml_env: Path) -> None:
    carregamento_id = _seed_carregamento([CHAVE])
    export = XmlExportService(documento_xml_service=DocumentoXmlService(storage_dir=xml_env))
    result = export.collect_xmls_for_carregamento(carregamento_id)
    assert result.entries == ()
    assert result.missing_nfs


def test_same_xml_shared_across_carregamentos(xml_env: Path) -> None:
    service = DocumentoXmlService(storage_dir=xml_env)
    service.persist_raw_xml(_sample_xml(CHAVE), original_filename="NF1426799.xml")

    first_id = _seed_carregamento([CHAVE])
    second_id = _seed_carregamento([CHAVE])
    assert first_id != second_id

    export_a = XmlExportService(documento_xml_service=service).collect_xmls_for_carregamento(first_id)
    export_b = XmlExportService(documento_xml_service=service).collect_xmls_for_carregamento(second_id)
    assert len(export_a.entries) == 1
    assert len(export_b.entries) == 1
    assert export_a.entries[0].conteudo == export_b.entries[0].conteudo
    assert len(list(xml_env.glob("*.xml"))) == 1


def test_zip_pdf_and_xml() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=True,
        entrega_selected=True,
        xml_selected=True,
        xml_entries=[("NF1426799.xml", b"<xml/>"), ("NF1426804.xml", b"<xml2/>")],
        numero_carga="000019",
    )
    assert name == "Carregamento_000019.zip"
    assert mime == "application/zip"
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        names = archive.namelist()
        assert "Minuta.pdf" in names
        assert "Romaneio.pdf" in names
        assert "XML/NF1426799.xml" in names
        assert "XML/NF1426804.xml" in names


def test_zip_only_xmls() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=None,
        entrega_pdf_bytes=None,
        carregamento_selected=False,
        entrega_selected=False,
        xml_selected=True,
        xml_entries=[("NF1426799.xml", b"<xml/>")],
        numero_carga="000019",
    )
    assert name.endswith(".zip")
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert archive.namelist() == ["XML/NF1426799.xml"]


def test_zip_only_pdfs_legacy_names() -> None:
    payload, name, mime, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=b"%PDF-romaneio%",
        carregamento_selected=True,
        entrega_selected=True,
        xml_selected=False,
        xml_entries=None,
        numero_carga="000005",
    )
    with zipfile.ZipFile(BytesIO(payload)) as archive:
        assert "minuta_carregamento.pdf" in archive.namelist()
        assert "minuta_entrega.pdf" in archive.namelist()


def test_missing_xml_warning_only() -> None:
    _, _, _, message = build_documentos_download_package(
        carregamento_pdf_bytes=b"%PDF-minuta%",
        entrega_pdf_bytes=None,
        carregamento_selected=True,
        entrega_selected=False,
        xml_selected=True,
        xml_entries=[],
        numero_carga="000019",
    )
    assert XML_MISSING_WARNING in message


def test_batch_documental_performance_50_xmls(xml_env: Path) -> None:
    import time

    service = DocumentoXmlService(storage_dir=xml_env)
    items: list[XmlDocumentalItem] = []
    for index in range(50):
        chave = f"{CHAVE[:39]}{index:05d}"
        payload = _sample_xml(chave, str(1426000 + index))
        file_hash = hashlib.sha256(payload).hexdigest()
        items.append(
            XmlDocumentalItem(
                file_bytes=payload,
                hash_sha256=file_hash,
                original_filename=f"NF{1426000 + index}.xml",
                chave_nfe=chave,
            )
        )
    started = time.perf_counter()
    result = service.persist_raw_xml_batch(items, usuario_id=1)
    elapsed = time.perf_counter() - started
    assert result.saved == 50
    assert elapsed < 5.0, f"Persistencia documental de 50 XMLs demorou {elapsed:.2f}s"
    print(f"persistencia documental 50 XMLs: {elapsed:.3f}s ({result.elapsed_ms:.1f} ms interno)")
