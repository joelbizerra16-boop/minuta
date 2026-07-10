"""Testes P1.1 — persistencia operacional XML via upsert delta."""

from __future__ import annotations

from pathlib import Path

import pytest
from sqlalchemy import event, select
from sqlalchemy.orm import Session

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine, get_session_factory
from infrastructure.models.nota_fiscal import NotaFiscalORM
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
    }


@pytest.fixture
def xml_repo(tmp_path: Path) -> SqlXmlRecordRepository:
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    db_path = tmp_path / "xml_p11.db"
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


def test_upsert_inserts_only_delta_without_touching_existing(xml_repo: SqlXmlRecordRepository) -> None:
    seed = [_sample_record(f"{index:06d}") for index in range(1, 11)]
    xml_repo.upsert_records(seed)
    assert xml_repo.count_records() == 10

    untouched = seed[0]
    delta = [_sample_record("000011"), _sample_record("000012")]
    xml_repo.upsert_records(delta)
    assert xml_repo.count_records() == 12

    with get_session_factory()() as session:
        row = session.scalars(
            select(NotaFiscalORM).where(NotaFiscalORM.numero_nf == untouched["NF"])
        ).one()
        assert row.destinatario == "Cliente Teste"


def test_upsert_updates_existing_record(xml_repo: SqlXmlRecordRepository) -> None:
    original = _sample_record("000100")
    xml_repo.upsert_records([original])

    updated = dict(original)
    updated["Destinatario"] = "Cliente Atualizado"
    updated["Items"] = [
        {"cProd": "002", "Descricao": "Produto Novo", "Qtd": 5.0, "Unidade": "UN", "Peso": 2.0}
    ]
    xml_repo.upsert_records([updated])

    records = xml_repo.list_all_records()
    match = next(item for item in records if item["NF"] == "000100")
    assert match["Destinatario"] == "Cliente Atualizado"
    assert match["Items"][0]["cProd"] == "002"


def test_upsert_single_flush_and_single_commit(xml_repo: SqlXmlRecordRepository) -> None:
    engine = get_engine()
    stats = {"flush": 0, "commit": 0, "execute": 0}

    @event.listens_for(Session, "after_flush")
    def _after_flush(session, flush_context) -> None:
        stats["flush"] += 1

    @event.listens_for(engine, "commit")
    def _on_commit(conn) -> None:
        stats["commit"] += 1

    @event.listens_for(engine, "before_cursor_execute")
    def _before_execute(conn, cursor, statement, parameters, context, executemany) -> None:
        stats["execute"] += 1

    records = [_sample_record(f"{index:06d}") for index in range(1, 6)]
    xml_repo.upsert_records(records)

    assert stats["flush"] == 1
    assert stats["commit"] == 1
    assert stats["execute"] < 30


def test_replace_all_still_available_for_cleanup_paths(xml_repo: SqlXmlRecordRepository) -> None:
    xml_repo.upsert_records([_sample_record("000200")])
    xml_repo.replace_all_records([])
    assert xml_repo.count_records() == 0
