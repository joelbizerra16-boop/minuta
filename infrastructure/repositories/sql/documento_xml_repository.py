from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.repositories.documento_xml_repository import DocumentoXmlRecord, DocumentoXmlRepository


class SqlDocumentoXmlRepository(DocumentoXmlRepository):
    def __init__(self, session: Session) -> None:
        self._session = session

    def get_by_chave(self, chave_nfe: str) -> DocumentoXmlRecord | None:
        if not chave_nfe:
            return None
        row = self._session.scalars(
            select(DocumentoXmlORM).where(
                DocumentoXmlORM.chave_nfe == chave_nfe,
                DocumentoXmlORM.ativo.is_(True),
            )
        ).first()
        return self._to_record(row) if row else None

    def get_by_numero_nf(self, numero_nf: str) -> DocumentoXmlRecord | None:
        if not numero_nf:
            return None
        row = self._session.scalars(
            select(DocumentoXmlORM)
            .where(
                DocumentoXmlORM.numero_nf == numero_nf,
                DocumentoXmlORM.ativo.is_(True),
            )
            .order_by(DocumentoXmlORM.id.desc())
        ).first()
        return self._to_record(row) if row else None

    def list_by_chaves(self, chaves: list[str]) -> dict[str, DocumentoXmlRecord]:
        normalized = [str(chave or "").strip() for chave in chaves if str(chave or "").strip()]
        if not normalized:
            return {}
        rows = self._session.scalars(
            select(DocumentoXmlORM).where(
                DocumentoXmlORM.chave_nfe.in_(normalized),
                DocumentoXmlORM.ativo.is_(True),
            )
        ).all()
        return {row.chave_nfe: self._to_record(row) for row in rows}

    def list_by_numeros_nf(self, numeros_nf: list[str]) -> dict[str, DocumentoXmlRecord]:
        normalized = [str(numero or "").strip() for numero in numeros_nf if str(numero or "").strip()]
        if not normalized:
            return {}
        rows = self._session.scalars(
            select(DocumentoXmlORM)
            .where(
                DocumentoXmlORM.numero_nf.in_(normalized),
                DocumentoXmlORM.ativo.is_(True),
            )
            .order_by(DocumentoXmlORM.id.desc())
        ).all()
        result: dict[str, DocumentoXmlRecord] = {}
        for row in rows:
            if row.numero_nf not in result:
                result[row.numero_nf] = self._to_record(row)
        return result

    def save(self, record: DocumentoXmlRecord) -> DocumentoXmlRecord:
        if record.id > 0:
            row = self._session.get(DocumentoXmlORM, record.id)
            if row is None:
                raise ValueError(f"DocumentoXml {record.id} nao encontrado.")
        else:
            existing = self._session.scalars(
                select(DocumentoXmlORM).where(DocumentoXmlORM.chave_nfe == record.chave_nfe)
            ).first()
            row = existing or DocumentoXmlORM()
            if existing is None:
                self._session.add(row)

        row.chave_nfe = record.chave_nfe
        row.numero_nf = record.numero_nf
        row.nome_arquivo = record.nome_arquivo
        row.caminho_arquivo = record.caminho_arquivo
        row.hash_sha256 = record.hash_sha256
        row.tamanho = int(record.tamanho)
        row.usuario_id = record.usuario_id
        row.ativo = bool(record.ativo)
        return self._to_record(row)

    @staticmethod
    def _to_record(row: DocumentoXmlORM) -> DocumentoXmlRecord:
        return DocumentoXmlRecord(
            id=int(row.id or 0),
            chave_nfe=row.chave_nfe,
            numero_nf=row.numero_nf,
            nome_arquivo=row.nome_arquivo,
            caminho_arquivo=row.caminho_arquivo,
            hash_sha256=row.hash_sha256,
            tamanho=int(row.tamanho),
            usuario_id=int(row.usuario_id) if row.usuario_id is not None else None,
            data_importacao=row.data_importacao,
            ativo=bool(row.ativo),
        )
