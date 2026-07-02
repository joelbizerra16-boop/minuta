from __future__ import annotations

import json

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.repositories.evento_auditoria_repository import EventoAuditoriaRecord, EventoAuditoriaRepository


class SqlEventoAuditoriaRepository(EventoAuditoriaRepository):
    def __init__(self, session: Session) -> None:
        self._session = session

    def append(self, evento: EventoAuditoriaRecord) -> EventoAuditoriaRecord:
        row = EventoAuditoriaORM(
            usuario_id=evento.usuario_id,
            categoria=evento.categoria,
            evento=evento.evento,
            entidade_tipo=evento.entidade_tipo,
            entidade_id=evento.entidade_id,
            descricao=evento.descricao,
            metadados_json=evento.metadados_json,
            ip_origem=evento.ip_origem,
        )
        self._session.add(row)
        self._session.flush()
        return self._to_record(row)

    def list_by_entidade(self, entidade_tipo: str, entidade_id: int) -> list[EventoAuditoriaRecord]:
        stmt = (
            select(EventoAuditoriaORM)
            .where(
                EventoAuditoriaORM.entidade_tipo == entidade_tipo,
                EventoAuditoriaORM.entidade_id == entidade_id,
            )
            .order_by(EventoAuditoriaORM.criado_em, EventoAuditoriaORM.id)
        )
        return [self._to_record(row) for row in self._session.scalars(stmt).all()]

    def list_by_usuario(self, usuario_id: int, limit: int = 100) -> list[EventoAuditoriaRecord]:
        stmt = (
            select(EventoAuditoriaORM)
            .where(EventoAuditoriaORM.usuario_id == usuario_id)
            .order_by(EventoAuditoriaORM.criado_em.desc(), EventoAuditoriaORM.id.desc())
            .limit(limit)
        )
        return [self._to_record(row) for row in self._session.scalars(stmt).all()]

    @staticmethod
    def build_metadados(**payload: object) -> str:
        return json.dumps(payload, ensure_ascii=False, sort_keys=True)

    @staticmethod
    def _to_record(row: EventoAuditoriaORM) -> EventoAuditoriaRecord:
        return EventoAuditoriaRecord(
            id=int(row.id),
            categoria=row.categoria,
            evento=row.evento,
            usuario_id=int(row.usuario_id) if row.usuario_id is not None else None,
            entidade_tipo=row.entidade_tipo,
            entidade_id=int(row.entidade_id) if row.entidade_id is not None else None,
            descricao=row.descricao,
            metadados_json=row.metadados_json,
            ip_origem=row.ip_origem,
            criado_em=row.criado_em,
        )
