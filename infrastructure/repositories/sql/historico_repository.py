from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.repositories.historico_repository import HistoricoRecord, HistoricoRepository


class SqlHistoricoRepository(HistoricoRepository):
    def __init__(self, session: Session) -> None:
        self._session = session

    def list_by_carregamento(self, carregamento_id: int) -> list[HistoricoRecord]:
        stmt = (
            select(HistoricoOperacionalORM)
            .where(HistoricoOperacionalORM.carregamento_id == carregamento_id)
            .order_by(HistoricoOperacionalORM.criado_em, HistoricoOperacionalORM.id)
        )
        return [self._to_record(row) for row in self._session.scalars(stmt).all()]

    def append(self, historico: HistoricoRecord) -> HistoricoRecord:
        row = HistoricoOperacionalORM(
            carregamento_id=historico.carregamento_id,
            usuario_id=historico.usuario_id,
            evento=historico.evento,
            descricao=historico.descricao,
        )
        self._session.add(row)
        self._session.flush()
        return self._to_record(row)

    @staticmethod
    def _to_record(row: HistoricoOperacionalORM) -> HistoricoRecord:
        return HistoricoRecord(
            id=int(row.id),
            carregamento_id=int(row.carregamento_id),
            usuario_id=int(row.usuario_id),
            evento=row.evento,
            descricao=row.descricao,
            criado_em=row.criado_em,
        )
