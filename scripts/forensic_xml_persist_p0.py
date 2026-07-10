#!/usr/bin/env python3
"""Forense P0: importacao de 1 XML novo com SQL audit e contagem antes/depois."""

from __future__ import annotations

import hashlib
import io
import json
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("MINUTA_SQL_AUDIT", "1")

from sqlalchemy import func, select, text

from core.bootstrap import configure_application_storage
from core.settings import get_settings
from infrastructure.database import get_engine
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.persistence.sql_audit import get_sql_audit_report, reset_sql_audit
from infrastructure.storage.xml_storage import SqlXmlRecordRepository


class _FakeUpload:
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def _build_proc_nfe_xml(*, chave: str, nf: str) -> bytes:
  nf_padded = nf.zfill(9)[-9:]
  chave_body = chave if len(chave) == 44 else ("4" * 44)
  xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<nfeProc xmlns="http://www.portalfiscal.inf.br/nfe" versao="4.00">
  <NFe xmlns="http://www.portalfiscal.inf.br/nfe">
    <infNFe Id="NFe{chave_body}" versao="4.00">
      <ide><nNF>{nf_padded}</nNF><dhEmi>2026-07-10T10:00:00-03:00</dhEmi></ide>
      <emit><xNome>EMITENTE TESTE</xNome></emit>
      <dest><xNome>DESTINATARIO TESTE</xNome><enderDest><xMun>Cidade</xMun><UF>SP</UF></enderDest></dest>
      <det nItem="1"><prod><cProd>001</cProd><xProd>Produto</xProd><qCom>1.0000</qCom><uCom>UN</uCom></prod></det>
      <total><ICMSTot><vNF>100.00</vNF></ICMSTot></total>
      <transp><vol><qVol>1</qVol><pesoL>1.000</pesoL></vol></transp>
    </infNFe>
  </NFe>
  <protNFe versao="4.00"><infProt><chNFe>{chave_body}</chNFe><xMotivo>Autorizado o uso da NF-e</xMotivo></infProt></protNFe>
</nfeProc>"""
  return xml.encode("utf-8")


def _table_counts() -> dict[str, int]:
    engine = get_engine()
    with engine.connect() as conn:
        return {
            "documento_xml": int(conn.execute(select(func.count()).select_from(DocumentoXmlORM)).scalar_one()),
            "nota_fiscal": int(conn.execute(select(func.count()).select_from(NotaFiscalORM)).scalar_one()),
            "item_nota_fiscal": int(conn.execute(select(func.count()).select_from(ItemNotaFiscalORM)).scalar_one()),
        }


def _sql_summary() -> dict[str, int]:
    report = get_sql_audit_report()
    summary: dict[str, int] = {}
    for entry in report.entries:
        op = str(entry.get("operation", "UNKNOWN")).upper()
        summary[op] = summary.get(op, 0) + 1
    return summary


def main() -> int:
    from app import import_xml_upload_batch, parse_xml_file, serialize_xml_record

    configure_application_storage()
    reset_sql_audit()

    unique_suffix = hashlib.sha256(str(time.time_ns()).encode()).hexdigest()[:9]
    chave = ("35260710" + unique_suffix + "0" * 27)[:44]
    nf = unique_suffix[-6:]

    before = _table_counts()
    xml_bytes = _build_proc_nfe_xml(chave=chave, nf=nf)
    upload = _FakeUpload(f"forensic_{nf}.xml", xml_bytes)

    parsed = parse_xml_file(upload)
    summary, issues = import_xml_upload_batch([upload])

    after = _table_counts()
    sql_ops = _sql_summary()

    exists = SqlXmlRecordRepository().list_all_records()
    from infrastructure.storage.xml_mapper import get_xml_storage_identity

    parsed_nf = str(parsed.get("NF", "") or parsed.get("nf_normalizada", "") or nf)
    target_identity = get_xml_storage_identity(
        {"NF": parsed_nf, "ChaveNFe": parsed.get("ChaveNFe", ""), "nf_normalizada": parsed_nf}
    )
    found = any(get_xml_storage_identity(item) == target_identity for item in exists)

    success = after["nota_fiscal"] > before["nota_fiscal"] and int(summary.get("novas", 0)) >= 1

    payload = {
        "chave_nfe": chave,
        "nf": nf,
        "counts_before": before,
        "counts_after": after,
        "summary": summary,
        "issues": issues,
        "sql_ops": sql_ops,
        "found_after_list_all": found,
        "persist_success": success,
        "parsed_chave": parsed.get("ChaveNFe"),
        "serialized_identity": serialize_xml_record(parsed).get("ChaveNFe"),
    }
    print(json.dumps(payload, indent=2, ensure_ascii=False))
    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())
