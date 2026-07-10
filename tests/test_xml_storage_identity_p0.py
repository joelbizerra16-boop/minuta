"""P0 — identidade de storage alinhada entre persist e ORM roundtrip."""

from __future__ import annotations

from pathlib import Path

import pytest

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.xml_mapper import get_xml_storage_identity, orm_to_record, record_to_orm
from infrastructure.storage.xml_storage import SqlXmlRecordRepository


@pytest.fixture
def xml_repo(tmp_path: Path) -> SqlXmlRecordRepository:
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "identity_p0.db"
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=tmp_path,
        pdf_storage_dir=tmp_path / "pdf",
        xml_storage_dir=tmp_path / "xml",
    )
    ensure_full_schema()
    yield SqlXmlRecordRepository()
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def test_roundtrip_preserves_storage_identity_for_hash_chave(xml_repo: SqlXmlRecordRepository) -> None:
    record = {
        "NF": "999001",
        "nf_normalizada": "999001",
        "ChaveNFe": "",
        "Destinatario": "Cliente",
        "StatusNF": "Autorizado",
        "Items": [],
        "Arquivo": "999001.xml",
    }
    xml_repo.upsert_records([record])
    loaded = xml_repo.list_all_records()[0]
    identity_loaded = get_xml_storage_identity(loaded)
    identity_source = get_xml_storage_identity(record)
    assert identity_loaded == identity_source
    assert identity_loaded != "999001"


def test_new_chave_inserts_when_not_in_database(xml_repo: SqlXmlRecordRepository) -> None:
    before = xml_repo.count_records()
    chave = "3" * 44
    record = {
        "NF": "888001",
        "nf_normalizada": "888001",
        "ChaveNFe": chave,
        "Destinatario": "Cliente Novo",
        "StatusNF": "Autorizado",
        "Items": [{"cProd": "1", "Descricao": "P", "Qtd": 1.0, "Unidade": "UN", "Peso": 1.0}],
        "Arquivo": "888001.xml",
    }
    xml_repo.upsert_records([record])
    assert xml_repo.count_records() == before + 1
    loaded = xml_repo.list_all_records()
    assert any(item["ChaveNFe"] == chave for item in loaded)


def test_legacy_orm_stripped_chave_caused_identity_mismatch() -> None:
    """Evidencia RCA: orm_to_record legado descartava chave hash e quebrava deduplicacao."""
    import re

    from app import normalize_chave_nfe, normalize_nf
    from infrastructure.storage.xml_mapper import _resolve_chave_nfe, get_xml_storage_identity, orm_to_record, record_to_orm

    nf = "654321"
    row = record_to_orm({"NF": nf, "ChaveNFe": "", "Destinatario": "X", "Items": []})
    hash_chave = _resolve_chave_nfe({"NF": nf, "ChaveNFe": ""})
    assert row.chave_nfe == hash_chave

    legacy_chave_field = row.chave_nfe if re.fullmatch(r"\d{44}", row.chave_nfe or "") else ""
    legacy_identity = (
        normalize_chave_nfe(legacy_chave_field)
        if normalize_chave_nfe(legacy_chave_field)
        else normalize_nf(nf)
    )

    fixed_loaded = orm_to_record(row, [])
    canonical_loaded = get_xml_storage_identity(fixed_loaded)
    canonical_upload = get_xml_storage_identity({"NF": nf, "ChaveNFe": "", "nf_normalizada": nf})

    assert canonical_loaded == canonical_upload == hash_chave
    assert legacy_identity != canonical_loaded


def test_orm_to_record_exposes_persisted_chave() -> None:
    row = record_to_orm({"NF": "777001", "ChaveNFe": "", "Destinatario": "X", "Items": []})
    assert len(row.chave_nfe) == 44
    payload = orm_to_record(row, [])
    assert payload["ChaveNFe"] == row.chave_nfe
