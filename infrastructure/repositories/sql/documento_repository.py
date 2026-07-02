from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.documento import DocumentoORM
from infrastructure.repositories.documento_repository import DocumentoRecord, DocumentoRepository


class SqlDocumentoRepository(DocumentoRepository):
    def __init__(self, session: Session) -> None:
        self._session = session

    def get_by_id(self, documento_id: int) -> DocumentoRecord | None:
        row = self._session.get(DocumentoORM, documento_id)
        return self._to_record(row) if row else None

    def list_by_carregamento(self, carregamento_id: int) -> list[DocumentoRecord]:
        stmt = select(DocumentoORM).where(DocumentoORM.carregamento_id == carregamento_id).order_by(DocumentoORM.id)
        return [self._to_record(row) for row in self._session.scalars(stmt).all()]

    def save(self, documento: DocumentoRecord) -> DocumentoRecord:
        if documento.id > 0:
            row = self._session.get(DocumentoORM, documento.id)
            if row is None:
                raise ValueError(f"Documento {documento.id} nao encontrado.")
        else:
            row = DocumentoORM()
            self._session.add(row)

        row.carregamento_id = documento.carregamento_id
        row.usuario_id = documento.usuario_id
        row.tipo = documento.tipo
        row.caminho_arquivo = documento.caminho_arquivo
        row.nome_arquivo = documento.nome_arquivo
        row.hash_sha256 = documento.hash_sha256
        self._session.flush()
        return self._to_record(row)

    @staticmethod
    def _to_record(row: DocumentoORM) -> DocumentoRecord:
        return DocumentoRecord(
            id=int(row.id),
            carregamento_id=int(row.carregamento_id),
            usuario_id=int(row.usuario_id),
            tipo=row.tipo,
            caminho_arquivo=row.caminho_arquivo,
            nome_arquivo=row.nome_arquivo,
            hash_sha256=row.hash_sha256,
            criado_em=row.criado_em,
        )
