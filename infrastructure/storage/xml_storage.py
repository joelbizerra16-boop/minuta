from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal
from typing import Any

from sqlalchemy import delete, func, or_, select
from sqlalchemy.orm import Session, selectinload

from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.storage.xml_mapper import _parse_float, orm_to_record, record_to_items, record_to_orm
from infrastructure.unit_of_work import UnitOfWork


class SqlXmlRecordRepository:
    def list_all_records(self) -> list[dict[str, Any]]:
        with UnitOfWork() as uow:
            stmt = (
                select(NotaFiscalORM)
                .options(selectinload(NotaFiscalORM.itens))
                .order_by(NotaFiscalORM.numero_nf, NotaFiscalORM.id)
            )
            rows = uow.session.scalars(stmt).all()
            return [orm_to_record(row, list(row.itens)) for row in rows]

    def list_records_by_identities(self, identities: set[str] | list[str]) -> list[dict[str, Any]]:
        """Carrega apenas NFs cujas identidades (chave 44 ou numero) estao no conjunto."""
        normalized = {str(value).strip() for value in identities if str(value or "").strip()}
        if not normalized:
            return []

        chaves = {value for value in normalized if len(value) == 44}
        numeros = {value for value in normalized if value not in chaves}
        conditions = []
        if chaves:
            conditions.append(NotaFiscalORM.chave_nfe.in_(sorted(chaves)))
        if numeros:
            conditions.append(NotaFiscalORM.numero_nf.in_(sorted(numeros)))
        if not conditions:
            return []

        with UnitOfWork() as uow:
            rows = uow.session.scalars(
                select(NotaFiscalORM)
                .options(selectinload(NotaFiscalORM.itens))
                .where(or_(*conditions))
                .order_by(NotaFiscalORM.numero_nf, NotaFiscalORM.id)
            ).all()
            return [orm_to_record(row, list(row.itens)) for row in rows]

    def replace_all_records(self, records: list[dict[str, Any]]) -> None:
        with UnitOfWork() as uow:
            existing_items = uow.session.scalars(select(ItemNotaFiscalORM)).all()
            for row in existing_items:
                uow.session.delete(row)
            existing = uow.session.scalars(select(NotaFiscalORM)).all()
            for row in existing:
                uow.session.delete(row)
            uow.session.flush()
            self._insert_records(uow.session, records)
        self._invalidate_signatures()

    def upsert_records(self, records: list[dict[str, Any]]) -> None:
        if not records:
            return

        with UnitOfWork() as uow:
            session = uow.session
            lookup = self._preload_existing_rows(session, records)
            for record in records:
                row = self._resolve_existing_row(record, lookup)
                if row is None:
                    row = record_to_orm(record)
                    session.add(row)
                else:
                    record_to_orm(record, row)
                if int(row.id or 0) > 0:
                    session.execute(
                        delete(ItemNotaFiscalORM).where(
                            ItemNotaFiscalORM.nota_fiscal_id == int(row.id)
                        )
                    )
                    row.itens.clear()
                    for item in record_to_items(record, int(row.id)):
                        session.add(item)
                else:
                    row.itens.clear()
                    SqlXmlRecordRepository._append_items_for_new_row(row, record)
            session.flush()
        self._invalidate_signatures()

    @staticmethod
    def _invalidate_signatures() -> None:
        from core.runtime_data_coherence import invalidate_data_signature_cache

        invalidate_data_signature_cache()

    @staticmethod
    def _append_items_for_new_row(row: NotaFiscalORM, record: dict[str, Any]) -> None:
        for index, item_data in enumerate(record.get("Items", []) or [], start=1):
            if not isinstance(item_data, dict):
                continue
            row.itens.append(
                ItemNotaFiscalORM(
                    sequencia=index,
                    codigo_produto=str(item_data.get("cProd", "") or "").strip() or "--",
                    descricao=str(item_data.get("Descricao", "") or "").strip() or "Sem descricao",
                    quantidade=Decimal(str(_parse_float(item_data.get("Qtd", 0)))),
                    unidade=str(item_data.get("Unidade", "") or "").strip() or None,
                    peso=Decimal(str(_parse_float(item_data.get("Peso", 0)))),
                )
            )

    def get_last_updated_at(self) -> datetime | None:
        with UnitOfWork() as uow:
            value = uow.session.scalar(select(func.max(NotaFiscalORM.atualizado_em)))
            if isinstance(value, datetime):
                return value
            return None

    def count_records(self) -> int:
        with UnitOfWork() as uow:
            return int(uow.session.scalar(select(func.count()).select_from(NotaFiscalORM)) or 0)

    @staticmethod
    def _record_keys(record: dict[str, Any]) -> tuple[str, str]:
        chave = str(record.get("ChaveNFe", "") or "").strip()
        numero = str(record.get("NF", "") or record.get("nf_normalizada", "") or "").strip()
        if not chave:
            chave = record_to_orm(record).chave_nfe
        return chave, numero

    def _preload_existing_rows(
        self,
        session: Session,
        records: list[dict[str, Any]],
    ) -> dict[str, NotaFiscalORM]:
        chaves: set[str] = set()
        numeros: set[str] = set()
        for record in records:
            chave, numero = self._record_keys(record)
            if chave:
                chaves.add(chave)
            if numero:
                numeros.add(numero)

        lookup: dict[str, NotaFiscalORM] = {}
        if chaves:
            rows = session.scalars(
                select(NotaFiscalORM)
                .options(selectinload(NotaFiscalORM.itens))
                .where(NotaFiscalORM.chave_nfe.in_(sorted(chaves)))
            ).all()
            for row in rows:
                lookup[f"chave:{row.chave_nfe}"] = row

        unresolved_numeros = {
            numero
            for record in records
            for chave, numero in [self._record_keys(record)]
            if numero and f"chave:{chave}" not in lookup
        }
        if unresolved_numeros:
            rows = session.scalars(
                select(NotaFiscalORM)
                .options(selectinload(NotaFiscalORM.itens))
                .where(NotaFiscalORM.numero_nf.in_(sorted(unresolved_numeros)))
                .order_by(NotaFiscalORM.id)
            ).all()
            seen_numeros: set[str] = set()
            for row in rows:
                if row.numero_nf in seen_numeros:
                    continue
                seen_numeros.add(row.numero_nf)
                lookup.setdefault(f"numero:{row.numero_nf}", row)

        return lookup

    @staticmethod
    def _resolve_existing_row(
        record: dict[str, Any],
        lookup: dict[str, NotaFiscalORM],
    ) -> NotaFiscalORM | None:
        chave, numero = SqlXmlRecordRepository._record_keys(record)
        if chave:
            found = lookup.get(f"chave:{chave}")
            if found is not None:
                return found
        if numero:
            return lookup.get(f"numero:{numero}")
        return None

    @staticmethod
    def _insert_records(session: Session, records: list[dict[str, Any]]) -> None:
        for record in records:
            row = record_to_orm(record)
            session.add(row)
            session.flush()
            for item in record_to_items(record, int(row.id)):
                session.add(item)
