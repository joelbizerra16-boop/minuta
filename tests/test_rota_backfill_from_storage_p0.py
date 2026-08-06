"""P0 — Processar Excel-only recupera rota a partir do XML em disco."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database, get_xml_storage_dir
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.xml_storage import SqlXmlRecordRepository


REFERENCE_XML = Path(
    r"c:\Users\joelb\Downloads\procNFe35260700846804000106550010014360591118433961.xml"
)
CHAVE = "35260700846804000106550010014360591118433961"
NF = "1436059"


@pytest.fixture
def xml_env(tmp_path: Path):
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "rota_backfill.db"
    xml_dir = tmp_path / "xml_storage"
    xml_dir.mkdir(parents=True, exist_ok=True)
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=tmp_path,
        pdf_storage_dir=tmp_path / "pdf",
        xml_storage_dir=xml_dir,
    )
    ensure_full_schema()
    yield SqlXmlRecordRepository(), xml_dir
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def test_enrich_from_disk_updates_nota_and_dataframe(xml_env) -> None:
    if not REFERENCE_XML.is_file():
        pytest.skip("XML de referência não disponível")

    from app import (
        apply_routes_from_xml_index,
        build_xml_index_from_records,
        enrich_xml_records_routes_from_storage,
        serialize_xml_record,
    )

    repo, xml_dir = xml_env
    assert get_xml_storage_dir() == xml_dir

    stale = serialize_xml_record(
        {
            "NF": NF,
            "nf_normalizada": NF,
            "ChaveNFe": CHAVE,
            "Destinatario": "PHP",
            "Municipio": "SBC",
            "UF": "SP",
            "StatusNF": "Autorizado o uso da NF-e",
            "Status": "Autorizado o uso da NF-e",
            "ROTA": "",
            "Items": [{"cProd": "1", "Descricao": "X", "Qtd": 1, "Unidade": "CX", "Peso": 1.0}],
            "Arquivo": "old.xml",
            "TipoXML": "normal",
            "ValorNF": 1.0,
            "PesoTotal": 1.0,
            "VolumeTotal": 1.0,
            "Data": "2026-07-31",
            "DataReferenciaISO": "2026-07-31T15:43:00-03:00",
            "Erro": False,
        }
    )
    assert stale["ROTA"] == "NÃO DEFINIDA"
    repo.upsert_records([stale])

    # Arquivo físico como em produção (xml_storage/{chave}.xml)
    (xml_dir / f"{CHAVE}.xml").write_bytes(REFERENCE_XML.read_bytes())

    loaded = repo.list_all_records()
    enriched, count = enrich_xml_records_routes_from_storage(loaded, only_nfs={NF}, persist=True)
    assert count == 1
    assert enriched[0]["ROTA"] == "ABCD"
    assert repo.list_all_records()[0]["ROTA"] == "ABCD"

    xml_index, _ = build_xml_index_from_records(enriched)
    dataframe = pd.DataFrame([{"NF": NF, "cProd": "1", "Descricao": "X"}])
    with_routes = apply_routes_from_xml_index(dataframe, xml_index)
    assert with_routes.iloc[0]["ROTA"] == "ABCD"


def test_enrich_skips_when_file_missing(xml_env) -> None:
    from app import enrich_xml_records_routes_from_storage, serialize_xml_record

    repo, _xml_dir = xml_env
    stale = serialize_xml_record(
        {
            "NF": NF,
            "nf_normalizada": NF,
            "ChaveNFe": CHAVE,
            "Destinatario": "PHP",
            "StatusNF": "Autorizado o uso da NF-e",
            "ROTA": "",
            "Items": [],
            "Arquivo": "old.xml",
            "TipoXML": "normal",
            "Erro": False,
        }
    )
    repo.upsert_records([stale])
    enriched, count = enrich_xml_records_routes_from_storage(repo.list_all_records(), persist=True)
    assert count == 0
    assert enriched[0]["ROTA"] == "NÃO DEFINIDA"


def test_enrich_route_from_database_xml_after_local_file_loss(xml_env) -> None:
    from app import enrich_xml_records_routes_from_storage, serialize_xml_record
    from infrastructure.services.documento_xml_service import DocumentoXmlService

    repo, xml_dir = xml_env
    payload = f"""<?xml version="1.0" encoding="UTF-8"?>
    <nfeProc>
      <NFe><infNFe><ide><nNF>{NF}</nNF></ide>
        <dest><xNome>CLIENTE</xNome><enderDest><xMun>MAIRIPORA</xMun><UF>SP</UF></enderDest></dest>
        <infAdic><infCpl>Cliente C015265 Rota MAIRIPORA\\nTrib aprox</infCpl></infAdic>
      </infNFe></NFe>
      <protNFe><infProt><chNFe>{CHAVE}</chNFe><cStat>100</cStat><xMotivo>Autorizado o uso da NF-e</xMotivo></infProt></protNFe>
    </nfeProc>""".encode("utf-8")

    stale = serialize_xml_record(
        {
            "NF": NF,
            "nf_normalizada": NF,
            "ChaveNFe": CHAVE,
            "Destinatario": "CLIENTE",
            "Municipio": "MAIRIPORA",
            "UF": "SP",
            "StatusNF": "Autorizado o uso da NF-e",
            "ROTA": "",
            "Items": [],
            "Arquivo": "legacy.xml",
            "TipoXML": "normal",
            "Erro": False,
        }
    )
    repo.upsert_records([stale])
    DocumentoXmlService(storage_dir=xml_dir).persist_raw_xml(
        payload,
        original_filename="legacy.xml",
        chave_nfe=CHAVE,
    )
    (xml_dir / f"{CHAVE}.xml").unlink()

    enriched, count = enrich_xml_records_routes_from_storage(repo.list_all_records(), persist=True)
    assert count == 1
    assert enriched[0]["ROTA"] == "MAIRIPORA"
    assert repo.list_all_records()[0]["ROTA"] == "MAIRIPORA"
