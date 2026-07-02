from __future__ import annotations

from decimal import Decimal

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.repositories.nota_fiscal_repository import (
    ItemNotaFiscalRecord,
    NotaFiscalRecord,
    NotaFiscalRepository,
)


class SqlNotaFiscalRepository(NotaFiscalRepository):
    def __init__(self, session: Session) -> None:
        self._session = session

    def get_by_id(self, nota_fiscal_id: int) -> NotaFiscalRecord | None:
        row = self._session.get(NotaFiscalORM, nota_fiscal_id)
        return self._to_record(row) if row else None

    def get_by_chave(self, chave_nfe: str) -> NotaFiscalRecord | None:
        stmt = select(NotaFiscalORM).where(NotaFiscalORM.chave_nfe == chave_nfe)
        row = self._session.scalars(stmt).first()
        return self._to_record(row) if row else None

    def list_all(self) -> list[NotaFiscalRecord]:
        stmt = select(NotaFiscalORM).order_by(NotaFiscalORM.id)
        return [self._to_record(row) for row in self._session.scalars(stmt).all()]

    def save(self, nota_fiscal: NotaFiscalRecord) -> NotaFiscalRecord:
        if nota_fiscal.id > 0:
            row = self._session.get(NotaFiscalORM, nota_fiscal.id)
            if row is None:
                raise ValueError(f"Nota fiscal {nota_fiscal.id} nao encontrada.")
        else:
            row = NotaFiscalORM()
            self._session.add(row)

        row.chave_nfe = nota_fiscal.chave_nfe
        row.numero_nf = nota_fiscal.numero_nf
        row.destinatario = nota_fiscal.destinatario
        row.status_nf = nota_fiscal.status_nf
        row.valor_total = nota_fiscal.valor_total
        row.peso_total = nota_fiscal.peso_total
        row.volume_total = nota_fiscal.volume_total
        row.emitente = nota_fiscal.emitente
        row.municipio = nota_fiscal.municipio
        row.uf = nota_fiscal.uf
        row.rota = nota_fiscal.rota
        row.tipo_xml = nota_fiscal.tipo_xml
        row.data_emissao = nota_fiscal.data_emissao
        row.data_referencia = nota_fiscal.data_referencia
        row.arquivo_origem = nota_fiscal.arquivo_origem
        row.destinatario_id = nota_fiscal.destinatario_id
        row.rota_id = nota_fiscal.rota_id
        self._session.flush()
        return self._to_record(row)

    def list_itens(self, nota_fiscal_id: int) -> list[ItemNotaFiscalRecord]:
        stmt = (
            select(ItemNotaFiscalORM)
            .where(ItemNotaFiscalORM.nota_fiscal_id == nota_fiscal_id)
            .order_by(ItemNotaFiscalORM.sequencia, ItemNotaFiscalORM.id)
        )
        return [self._to_item_record(row) for row in self._session.scalars(stmt).all()]

    def save_itens(self, nota_fiscal_id: int, itens: list[ItemNotaFiscalRecord]) -> list[ItemNotaFiscalRecord]:
        existing = self._session.scalars(
            select(ItemNotaFiscalORM).where(ItemNotaFiscalORM.nota_fiscal_id == nota_fiscal_id)
        ).all()
        for row in existing:
            self._session.delete(row)
        self._session.flush()

        saved: list[ItemNotaFiscalRecord] = []
        for item in itens:
            row = ItemNotaFiscalORM(
                nota_fiscal_id=nota_fiscal_id,
                sequencia=item.sequencia,
                codigo_produto=item.codigo_produto,
                descricao=item.descricao,
                quantidade=item.quantidade,
                unidade=item.unidade,
                peso=item.peso,
            )
            self._session.add(row)
            self._session.flush()
            saved.append(self._to_item_record(row))
        return saved

    @staticmethod
    def _to_record(row: NotaFiscalORM) -> NotaFiscalRecord:
        return NotaFiscalRecord(
            id=int(row.id),
            chave_nfe=row.chave_nfe,
            numero_nf=row.numero_nf,
            destinatario=row.destinatario,
            status_nf=row.status_nf,
            valor_total=Decimal(row.valor_total),
            peso_total=Decimal(row.peso_total),
            volume_total=Decimal(row.volume_total),
            emitente=row.emitente,
            municipio=row.municipio,
            uf=row.uf,
            rota=row.rota,
            tipo_xml=row.tipo_xml,
            data_emissao=row.data_emissao,
            data_referencia=row.data_referencia,
            arquivo_origem=row.arquivo_origem,
            destinatario_id=row.destinatario_id,
            rota_id=row.rota_id,
        )

    @staticmethod
    def _to_item_record(row: ItemNotaFiscalORM) -> ItemNotaFiscalRecord:
        return ItemNotaFiscalRecord(
            id=int(row.id),
            nota_fiscal_id=int(row.nota_fiscal_id),
            sequencia=int(row.sequencia),
            codigo_produto=row.codigo_produto,
            descricao=row.descricao,
            quantidade=Decimal(row.quantidade),
            peso=Decimal(row.peso),
            unidade=row.unidade,
        )
