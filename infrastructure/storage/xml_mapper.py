from __future__ import annotations

import hashlib
import re
from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Any

from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM


def _parse_float(value: object) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _parse_date(value: object) -> date | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw[:10], fmt).date()
        except ValueError:
            continue
    return None


def _parse_datetime(value: object) -> datetime | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    normalized = raw.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed


def _resolve_chave_nfe(record: dict[str, Any]) -> str:
    digits = re.sub(r"\D", "", str(record.get("ChaveNFe", "") or ""))
    if len(digits) == 44:
        return digits
    numero = str(record.get("NF", "") or record.get("nf_normalizada", "") or "").strip()
    digest = hashlib.sha256(f"NF:{numero}".encode("utf-8")).hexdigest()
    return digest[:44]


def get_xml_storage_identity(record: dict[str, Any]) -> str:
    """Identidade canonica para deduplicacao alinhada ao valor persistido em nota_fiscal.chave_nfe."""
    raw_chave = str(record.get("ChaveNFe", "") or "").strip()
    if raw_chave:
        digits = re.sub(r"\D", "", raw_chave)
        if len(digits) == 44:
            return digits
        if len(raw_chave) == 44:
            return raw_chave
    resolved = _resolve_chave_nfe(record)
    if resolved:
        return resolved
    return str(record.get("NF", "") or record.get("nf_normalizada", "") or "").strip()


def record_to_orm(record: dict[str, Any], row: NotaFiscalORM | None = None) -> NotaFiscalORM:
    target = row or NotaFiscalORM()
    target.chave_nfe = _resolve_chave_nfe(record)
    target.numero_nf = str(record.get("NF", "") or record.get("nf_normalizada", "") or "").strip()
    target.destinatario = str(record.get("Destinatario", "") or "").strip() or "Nao informado"
    target.municipio = str(record.get("Municipio", "") or "").strip() or None
    target.uf = str(record.get("UF", "") or "").strip() or None
    target.rota = str(record.get("ROTA", "") or "").strip() or None
    target.status_nf = str(record.get("StatusNF", record.get("Status", "")) or "").strip() or "Nao informado"
    target.tipo_xml = str(record.get("TipoXML", "normal") or "normal").strip()
    target.valor_total = Decimal(str(_parse_float(record.get("ValorNF", 0))))
    target.peso_total = Decimal(str(_parse_float(record.get("PesoTotal", 0))))
    target.volume_total = Decimal(str(_parse_float(record.get("VolumeTotal", 0))))
    target.arquivo_origem = str(record.get("Arquivo", "") or "").strip() or None
    target.data_emissao = _parse_date(record.get("Data", ""))
    target.data_referencia = _parse_datetime(
        record.get("DataReferenciaISO", "") or record.get("DataReferencia", "") or record.get("Data", "")
    )
    return target


def orm_to_record(row: NotaFiscalORM, itens: list[ItemNotaFiscalORM]) -> dict[str, Any]:
    items_payload = [
        {
            "cProd": item.codigo_produto,
            "Descricao": item.descricao,
            "Qtd": float(item.quantidade),
            "Unidade": item.unidade or "",
            "Peso": float(item.peso),
        }
        for item in itens
    ]
    data_ref = row.data_referencia.isoformat() if row.data_referencia else ""
    data_label = row.data_emissao.isoformat() if row.data_emissao else ""
    return {
        "NF": row.numero_nf,
        "nf_normalizada": row.numero_nf,
        "ChaveNFe": row.chave_nfe or "",
        "Data": data_label,
        "DataReferencia": data_ref,
        "DataReferenciaISO": data_ref,
        "Destinatario": row.destinatario,
        "Municipio": row.municipio or "",
        "UF": row.uf or "",
        "Status": row.status_nf,
        "StatusNF": row.status_nf,
        "ValorNF": float(row.valor_total),
        "VolumeTotal": float(row.volume_total),
        "PesoTotal": float(row.peso_total),
        "Items": items_payload,
        "Arquivo": row.arquivo_origem or "",
        "Erro": False,
        "TipoXML": row.tipo_xml or "normal",
        "ROTA": row.rota or "",
    }


def record_to_items(record: dict[str, Any], nota_fiscal_id: int) -> list[ItemNotaFiscalORM]:
    items: list[ItemNotaFiscalORM] = []
    for index, item in enumerate(record.get("Items", []) or [], start=1):
        if not isinstance(item, dict):
            continue
        items.append(
            ItemNotaFiscalORM(
                nota_fiscal_id=nota_fiscal_id,
                sequencia=index,
                codigo_produto=str(item.get("cProd", "") or "").strip() or "--",
                descricao=str(item.get("Descricao", "") or "").strip() or "Sem descricao",
                quantidade=Decimal(str(_parse_float(item.get("Qtd", 0)))),
                unidade=str(item.get("Unidade", "") or "").strip() or None,
                peso=Decimal(str(_parse_float(item.get("Peso", 0)))),
            )
        )
    return items
