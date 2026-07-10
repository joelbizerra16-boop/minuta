"""Benchmark P1.2 — persistencia documental XML: flush por registro vs flush unico."""

from __future__ import annotations

import hashlib
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _sample_xml(chave: str) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc><NFe><infNFe><ide><nNF>1426799</nNF></ide><chNFe>{chave}</chNFe></infNFe></NFe></nfeProc>""".encode(
        "utf-8"
    )


def _build_items(count: int) -> list:
    from infrastructure.services.documento_xml_service import XmlDocumentalItem

    base = "35260612345678901234550010000012345678901234"
    items = []
    for index in range(count):
        chave = f"{base[:39]}{index:05d}"
        payload = _sample_xml(chave)
        items.append(
            XmlDocumentalItem(
                file_bytes=payload,
                hash_sha256=hashlib.sha256(payload).hexdigest(),
                original_filename=f"NF{index}.xml",
                chave_nfe=chave,
            )
        )
    return items


def _measure_batch(*, batch_size: int) -> dict[str, float | int]:
    from sqlalchemy import delete, event
    from sqlalchemy.orm import Session

    from auth.bootstrap import configure_auth_storage
    from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
    from core.settings import get_settings
    from infrastructure.database import configure_database, get_engine
    from infrastructure.models.documento_xml import DocumentoXmlORM
    from infrastructure.schema import ensure_full_schema
    from infrastructure.services.documento_xml_service import DocumentoXmlService
    from infrastructure.unit_of_work import UnitOfWork

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        get_settings.cache_clear()
        reset_infrastructure_bootstrap_state()
        storage_dir = tmp_path / "xml_storage"
        configure_database(
            database_url=f"sqlite:///{(tmp_path / 'bench_p12.db').as_posix()}",
            data_root=tmp_path,
            pdf_storage_dir=tmp_path / "pdf",
            xml_storage_dir=storage_dir,
        )
        ensure_full_schema()
        configure_auth_storage(tmp_path)

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
            uow.session.execute(delete(DocumentoXmlORM))

        stats["flush"] = 0
        stats["commit"] = 0
        stats["execute"] = 0

        items = _build_items(batch_size)
        service = DocumentoXmlService(storage_dir=storage_dir)

        start = time.perf_counter()
        result = service.persist_raw_xml_batch(items, usuario_id=1)
        elapsed_ms = (time.perf_counter() - start) * 1000

        reset_infrastructure_bootstrap_state()

    return {
        "elapsed_ms": round(elapsed_ms, 2),
        "internal_ms": round(result.elapsed_ms, 2),
        "saved": result.saved,
        "flushes": stats["flush"],
        "commits": stats["commit"],
        "queries": stats["execute"],
    }


def main() -> int:
    batch_size = int(__import__("os").getenv("BENCH_DOC_XML_SIZE", "50"))
    metrics = _measure_batch(batch_size=batch_size)

    print("## Benchmark P1.2 — Persistencia documental XML")
    print()
    print(f"Lote: {batch_size} XMLs novos")
    print()
    print("| Metrica | Valor |")
    print("| --- | ---: |")
    for key, value in metrics.items():
        print(f"| {key} | {value} |")
    print()
    print("Esperado pos-P1.2: flushes=1, commits=1")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
