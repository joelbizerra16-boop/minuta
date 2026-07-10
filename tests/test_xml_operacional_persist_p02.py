"""P0.2 — decisao operacional de importacao XML usa PostgreSQL, nao cache."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest
from sqlalchemy import delete, func, select

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.schema import ensure_full_schema
from infrastructure.storage.xml_storage import SqlXmlRecordRepository


def _sample_record(nf: str, *, chave: str | None = None) -> dict[str, object]:
    resolved_chave = chave or f"{int(nf):044d}"
    return {
        "NF": nf,
        "nf_normalizada": nf,
        "ChaveNFe": resolved_chave,
        "Destinatario": "Cliente Teste",
        "Municipio": "Cidade",
        "UF": "SP",
        "StatusNF": "Autorizado",
        "ValorNF": 100.0,
        "PesoTotal": 10.0,
        "VolumeTotal": 1.0,
        "Items": [{"cProd": "001", "Descricao": "Produto", "Qtd": 2.0, "Unidade": "UN", "Peso": 1.0}],
        "Arquivo": f"{nf}.xml",
        "TipoXML": "normal",
        "ROTA": "R1",
        "Data": "2026-01-01",
        "DataReferenciaISO": "2026-01-01T00:00:00+00:00",
        "Erro": False,
    }


@pytest.fixture
def xml_db(tmp_path: Path) -> Path:
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "xml_p02.db"
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=tmp_path,
        pdf_storage_dir=tmp_path / "pdf",
        xml_storage_dir=tmp_path / "xml",
    )
    ensure_full_schema()
    yield db_path
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def _count_nota_fiscal() -> int:
    engine = get_engine()
    with engine.connect() as conn:
        return int(conn.execute(select(func.count()).select_from(NotaFiscalORM)).scalar_one())


def _wipe_nota_fiscal() -> None:
    engine = get_engine()
    with engine.begin() as conn:
        conn.execute(delete(ItemNotaFiscalORM))
        conn.execute(delete(NotaFiscalORM))


def test_persist_uses_repository_not_stale_cache(xml_db: Path) -> None:
    """Cenario RCA: cache quente + PG vazio → operacional deve classificar NOVA."""
    repo = SqlXmlRecordRepository()
    chave = "1" * 44
    repo.upsert_records([_sample_record("000001", chave=chave)])

    from app import carregar_xmls_processados_records, persist_xml_records

    path = str(xml_db.parent / "xmls_processados.json")
    carregar_xmls_processados_records(str(path))  # warm cache with 1 record
    _wipe_nota_fiscal()
    assert _count_nota_fiscal() == 0

    upload = _sample_record("000001", chave=chave)
    upload["Arquivo"] = "dup_retry.xml"
    summary, issues = persist_xml_records(
        [upload],
        {"total_arquivos": 1, "erros": 0, "duplicados_lote": 0},
        [],
    )

    assert summary["novas"] == 1
    assert summary["duplicados_armazenamento"] == 0
    assert summary["processados"] == 1
    assert _count_nota_fiscal() == 1
    assert not any("duplicado ou desatualizado" in issue for issue in issues)


def test_persist_classifies_true_duplicate_from_postgresql(xml_db: Path) -> None:
    repo = SqlXmlRecordRepository()
    chave = "2" * 44
    record = _sample_record("000002", chave=chave)
    repo.upsert_records([record])

    from app import persist_xml_records

    retry = dict(record)
    retry["Arquivo"] = "retry.xml"
    summary, _ = persist_xml_records(
        [retry],
        {"total_arquivos": 1, "erros": 0, "duplicados_lote": 0},
        [],
    )

    assert summary["novas"] == 0
    assert summary["duplicados_armazenamento"] == 1
    assert summary["processados"] == 0
    assert _count_nota_fiscal() == 1


def test_persist_calls_list_records_by_identities_not_carregar_cache(xml_db: Path) -> None:
    from app import persist_xml_records

    repo_calls: list[int] = []
    original_list = SqlXmlRecordRepository.list_records_by_identities

    def _spy_list(self: SqlXmlRecordRepository, identities) -> list[dict[str, object]]:
        repo_calls.append(len(identities))
        return original_list(self, identities)

    stale_payload = ([_sample_record("999999")], "")

    with patch.object(SqlXmlRecordRepository, "list_records_by_identities", _spy_list):
        with patch.object(SqlXmlRecordRepository, "list_all_records", side_effect=AssertionError("list_all_records nao deve ser chamado")):
            with patch("app.carregar_xmls_processados_json", return_value=stale_payload) as cached_loader:
                persist_xml_records(
                    [_sample_record("000010")],
                    {"total_arquivos": 1, "erros": 0, "duplicados_lote": 0},
                    [],
                )

    assert len(repo_calls) >= 1
    assert repo_calls[0] == 1
    cached_loader.assert_not_called()


def test_interface_loader_still_uses_cache(xml_db: Path) -> None:
    from app import carregar_xmls_processados_records

    repo = SqlXmlRecordRepository()
    repo.upsert_records([_sample_record("000020")])

    path = str(xml_db.parent / "xmls.json")
    first = carregar_xmls_processados_records(path)
    second = carregar_xmls_processados_records(path)

    assert first == second
    assert len(first[0]) == 1


def test_bulk_import_twenty_new_records(xml_db: Path) -> None:
    from app import persist_xml_records

    records = [_sample_record(f"{index:06d}") for index in range(1, 21)]
    summary, _ = persist_xml_records(
        records,
        {"total_arquivos": 20, "erros": 0, "duplicados_lote": 0},
        [],
    )

    assert summary["novas"] == 20
    assert summary["processados"] == 20
    assert _count_nota_fiscal() == 20
