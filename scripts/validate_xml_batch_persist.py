"""Validacao rapida da persistencia documental em lote (sem Streamlit)."""
from __future__ import annotations

import hashlib
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from auth.bootstrap import configure_auth_storage
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.schema import ensure_full_schema
from infrastructure.services.documento_xml_service import DocumentoXmlService, XmlDocumentalItem

CHAVE = "35260612345678901234550010000012345678901234"


def sample_xml(chave: str) -> bytes:
    return f'<?xml version="1.0"?><nfeProc><NFe><infNFe><chNFe>{chave}</chNFe></infNFe></NFe></nfeProc>'.encode()


def main() -> None:
    get_settings.cache_clear()
    root = Path(tempfile.mkdtemp())
    db_path = (root / "t.db").as_posix()
    configure_database(
        database_url=f"sqlite:///{db_path}",
        data_root=root,
        pdf_storage_dir=root / "doc",
        xml_storage_dir=root / "xml",
    )
    ensure_full_schema()
    configure_auth_storage(root)

    service = DocumentoXmlService(storage_dir=root / "xml")
    items: list[XmlDocumentalItem] = []
    for index in range(50):
        chave = f"{CHAVE[:39]}{index:05d}"
        payload = sample_xml(chave)
        items.append(
            XmlDocumentalItem(
                file_bytes=payload,
                hash_sha256=hashlib.sha256(payload).hexdigest(),
                original_filename=f"NF{index}.xml",
                chave_nfe=chave,
            )
        )

    started = time.perf_counter()
    result = service.persist_raw_xml_batch(items, usuario_id=1)
    elapsed = time.perf_counter() - started
    print(f"saved={result.saved} reused={result.reused} elapsed_s={elapsed:.3f} internal_ms={result.elapsed_ms:.1f}")
    assert result.saved == 50, f"esperado 50 salvos, obteve {result.saved}"
    assert len(list((root / "xml").glob("*.xml"))) == 50
    print("VALIDACAO OK")


if __name__ == "__main__":
    main()
