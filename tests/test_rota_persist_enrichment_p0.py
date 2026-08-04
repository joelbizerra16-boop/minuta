"""P0 — reimport deve enriquecer rota vazia/NÃO DEFINIDA com rota válida do XML."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.xml_storage import SqlXmlRecordRepository
from utils.rota_xml import (
    UNDEFINED_ROUTE_LABEL,
    has_concrete_route,
    is_undefined_route,
    should_enrich_xml_route,
)


@pytest.fixture
def xml_env(tmp_path: Path) -> SqlXmlRecordRepository:
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "rota_enrich.db"
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


def _base_record(*, rota: str, arquivo: str = "nf.xml") -> dict[str, object]:
    return {
        "NF": "1436059",
        "nf_normalizada": "1436059",
        "ChaveNFe": "35260700846804000106550010014360591118433961",
        "Destinatario": "PHP COM ATACADISTA",
        "Municipio": "Sao Bernardo do Campo",
        "UF": "SP",
        "StatusNF": "Autorizado o uso da NF-e",
        "Status": "Autorizado o uso da NF-e",
        "ROTA": rota,
        "Items": [{"cProd": "123075", "Descricao": "MOBIL", "Qtd": 200, "Unidade": "CX", "Peso": 4620.0}],
        "Arquivo": arquivo,
        "TipoXML": "normal",
        "ValorNF": 127960.0,
        "PesoTotal": 4620.0,
        "VolumeTotal": 200.0,
        "Data": "2026-07-31",
        "DataReferencia": "2026-07-31T15:43:00-03:00",
        "DataReferenciaISO": "2026-07-31T15:43:00-03:00",
        "Erro": False,
    }


def test_should_enrich_xml_route_contract() -> None:
    assert is_undefined_route("")
    assert is_undefined_route(UNDEFINED_ROUTE_LABEL)
    assert is_undefined_route(None)
    assert has_concrete_route("ABCD")
    assert should_enrich_xml_route({"ROTA": ""}, {"ROTA": "ABCD"})
    assert should_enrich_xml_route({"ROTA": UNDEFINED_ROUTE_LABEL}, {"ROTA": "ABCD"})
    assert not should_enrich_xml_route({"ROTA": "R01"}, {"ROTA": "ABCD"})
    assert not should_enrich_xml_route({"ROTA": ""}, {"ROTA": UNDEFINED_ROUTE_LABEL})


def test_persist_enriches_undefined_route_on_reimport(xml_env: SqlXmlRecordRepository) -> None:
    from app import apply_routes_from_xml_index, build_xml_index_from_records, persist_xml_records, serialize_xml_record

    old = serialize_xml_record(_base_record(rota="", arquivo="old.xml"))
    assert old["ROTA"] == UNDEFINED_ROUTE_LABEL
    xml_env.upsert_records([old])
    assert xml_env.list_all_records()[0]["ROTA"] in {"", UNDEFINED_ROUTE_LABEL}

    new = serialize_xml_record(_base_record(rota="ABCD", arquivo="new.xml"))
    summary, issues = persist_xml_records(
        [new],
        {"total_arquivos": 1, "erros": 0, "duplicados_lote": 0},
        [],
    )
    assert summary["atualizadas"] == 1
    assert summary["duplicados_armazenamento"] == 0
    assert any("rota" in issue.lower() for issue in issues)

    loaded = xml_env.list_all_records()[0]
    assert loaded["ROTA"] == "ABCD"

    xml_index, _ = build_xml_index_from_records([loaded])
    dataframe = pd.DataFrame([{"NF": "1436059", "Descricao": "x", "cProd": "1"}])
    with_routes = apply_routes_from_xml_index(dataframe, xml_index)
    assert with_routes.iloc[0]["ROTA"] == "ABCD"


def test_persist_does_not_overwrite_existing_concrete_route(xml_env: SqlXmlRecordRepository) -> None:
    from app import persist_xml_records, serialize_xml_record

    current = serialize_xml_record(_base_record(rota="R01", arquivo="a.xml"))
    xml_env.upsert_records([current])

    challenger = serialize_xml_record(_base_record(rota="ABCD", arquivo="b.xml"))
    summary, _ = persist_xml_records(
        [challenger],
        {"total_arquivos": 1, "erros": 0, "duplicados_lote": 0},
        [],
    )
    assert summary["atualizadas"] == 0
    assert summary["duplicados_armazenamento"] == 1
    assert xml_env.list_all_records()[0]["ROTA"] == "R01"


def test_batch_dedup_enriches_route_from_later_file() -> None:
    from app import parse_xml_upload_batch

    class _Uploaded:
        def __init__(self, name: str, payload: bytes) -> None:
            self.name = name
            self._payload = payload

        def getvalue(self) -> bytes:
            return self._payload

    xml_path = Path(r"c:\Users\joelb\Downloads\procNFe35260700846804000106550010014360591118433961.xml")
    if not xml_path.is_file():
        pytest.skip("XML de referência não disponível")

    payload = xml_path.read_bytes()
    # Simula lote com o mesmo XML duas vezes — a rota deve permanecer ABCD
    records, summary, _ = parse_xml_upload_batch(
        [_Uploaded("a.xml", payload), _Uploaded("b.xml", payload)]
    )
    assert len(records) == 1
    assert records[0]["ROTA"] == "ABCD"
    assert records[0]["NF"] == "1436059"
