from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.storage.xml_mapper import orm_to_record, record_to_items, record_to_orm
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

    def upsert_records(self, records: list[dict[str, Any]]) -> None:
        with UnitOfWork() as uow:
            for record in records:
                chave = str(record.get("ChaveNFe", "") or "").strip()
                numero = str(record.get("NF", "") or record.get("nf_normalizada", "") or "").strip()
                row = None
                if chave:
                    row = uow.session.scalars(
                        select(NotaFiscalORM).where(NotaFiscalORM.chave_nfe == chave)
                    ).first()
                if row is None and numero:
                    row = uow.session.scalars(
                        select(NotaFiscalORM).where(NotaFiscalORM.numero_nf == numero)
                    ).first()
                if row is None:
                    row = record_to_orm(record)
                    uow.session.add(row)
                    uow.session.flush()
                else:
                    record_to_orm(record, row)
                for item in list(row.itens):
                    uow.session.delete(item)
                uow.session.flush()
                for item in record_to_items(record, int(row.id)):
                    uow.session.add(item)

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
    def _insert_records(session: Session, records: list[dict[str, Any]]) -> None:
        for record in records:
            row = record_to_orm(record)
            session.add(row)
            session.flush()
            for item in record_to_items(record, int(row.id)):
                session.add(item)
