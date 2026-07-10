"""Benchmark P1.1 — importacao operacional XML: replace_all vs upsert delta."""

from __future__ import annotations

import copy
import os
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _sample_record(index: int) -> dict[str, object]:
    nf = f"{index:06d}"
    return {
        "NF": nf,
        "nf_normalizada": nf,
        "ChaveNFe": f"{index:044d}",
        "Destinatario": "Cliente Benchmark",
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


def _measure(repo_method, base_records: list[dict[str, object]], delta: list[dict[str, object]]) -> dict[str, float | int]:
    from sqlalchemy import delete, event
    from sqlalchemy.orm import Session

    from infrastructure.database import get_engine
    from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
    from infrastructure.storage.xml_storage import SqlXmlRecordRepository
    from infrastructure.unit_of_work import UnitOfWork

    repo = SqlXmlRecordRepository()
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

    with UnitOfWork() as uow:
        uow.session.execute(delete(ItemNotaFiscalORM))
        uow.session.execute(delete(NotaFiscalORM))

    repo.upsert_records(base_records)
    stats["flush"] = 0
    stats["commit"] = 0
    stats["execute"] = 0

    merged = {str(item["ChaveNFe"]): item for item in base_records}
    for item in delta:
        merged[str(item["ChaveNFe"])] = item
    full_payload = list(merged.values())

    start = time.perf_counter()
    if repo_method == "replace_all":
        repo.replace_all_records(full_payload)
    else:
        repo.upsert_records(delta)
    elapsed_ms = (time.perf_counter() - start) * 1000
    return {
        "elapsed_ms": round(elapsed_ms, 2),
        "flushes": stats["flush"],
        "commits": stats["commit"],
        "queries": stats["execute"],
        "records_processed": len(delta),
        "full_table_records": len(full_payload),
    }


def _setup_db(db_path: Path) -> None:
    from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
    from core.settings import get_settings
    from infrastructure.database import configure_database
    from infrastructure.schema import ensure_full_schema

    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=db_path.parent,
        pdf_storage_dir=db_path.parent / "pdf",
        xml_storage_dir=db_path.parent / "xml",
    )
    ensure_full_schema()


def main() -> int:
    base_size = int(os.getenv("BENCH_XML_BASE_SIZE", "200"))
    delta_size = int(os.getenv("BENCH_XML_DELTA_SIZE", "20"))

    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "bench_xml_p11.db"
        _setup_db(db_path)

        base_records = [_sample_record(index) for index in range(1, base_size + 1)]
        delta = [_sample_record(base_size + index) for index in range(1, delta_size + 1)]

        before = _measure("replace_all", base_records, delta)

        _setup_db(db_path)
        base_copy = copy.deepcopy(base_records)
        delta_copy = copy.deepcopy(delta)
        after = _measure("upsert", base_copy, delta_copy)

        from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state

        reset_infrastructure_bootstrap_state()

        print("## Benchmark P1.1 — Importacao operacional XML")
        print()
        print(f"Base existente: {base_size} NF | Delta importado: {delta_size} NF")
        print()
        print("| Estrategia | Tempo (ms) | Queries | Flushes | Commits | Registros persistidos |")
        print("| --- | ---: | ---: | ---: | ---: | ---: |")
        print(
            f"| replace_all (antes) | {before['elapsed_ms']} | {before['queries']} | {before['flushes']} | "
            f"{before['commits']} | {before['full_table_records']} |"
        )
        print(
            f"| upsert delta (depois) | {after['elapsed_ms']} | {after['queries']} | {after['flushes']} | "
            f"{after['commits']} | {after['records_processed']} |"
        )
        speedup = before["elapsed_ms"] / after["elapsed_ms"] if after["elapsed_ms"] else 0
        print()
        print(f"Speedup local: {speedup:.1f}x")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
