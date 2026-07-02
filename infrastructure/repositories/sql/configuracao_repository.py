from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.configuracao import ConfiguracaoORM
from infrastructure.models.constants import CONFIG_TIPO_JSON
from infrastructure.repositories.configuracao_repository import ConfiguracaoRecord, ConfiguracaoRepository
from infrastructure.unit_of_work import UnitOfWork


class SqlConfiguracaoRepository(ConfiguracaoRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    def get_by_chave(self, chave: str) -> ConfiguracaoRecord | None:
        with self._uow() as uow:
            stmt = select(ConfiguracaoORM).where(ConfiguracaoORM.chave == chave)
            row = uow.session.scalars(stmt).first()
            return self._to_record(row) if row else None

    def list_all(self) -> list[ConfiguracaoRecord]:
        with self._uow() as uow:
            stmt = select(ConfiguracaoORM).order_by(ConfiguracaoORM.chave)
            return [self._to_record(row) for row in uow.session.scalars(stmt).all()]

    def save(self, configuracao: ConfiguracaoRecord) -> ConfiguracaoRecord:
        with self._uow() as uow:
            stmt = select(ConfiguracaoORM).where(ConfiguracaoORM.chave == configuracao.chave)
            row = uow.session.scalars(stmt).first()
            if row is None:
                row = ConfiguracaoORM(
                    chave=configuracao.chave,
                    valor=configuracao.valor,
                    categoria=configuracao.categoria,
                    tipo_valor=configuracao.tipo_valor,
                    descricao=configuracao.descricao,
                    atualizado_por_usuario_id=configuracao.atualizado_por_usuario_id,
                )
                uow.session.add(row)
            else:
                row.valor = configuracao.valor
                if configuracao.categoria:
                    row.categoria = configuracao.categoria
                if configuracao.tipo_valor:
                    row.tipo_valor = configuracao.tipo_valor
                row.descricao = configuracao.descricao
                row.atualizado_por_usuario_id = configuracao.atualizado_por_usuario_id
            uow.session.flush()
            return self._to_record(row)

    def delete(self, chave: str) -> bool:
        with self._uow() as uow:
            stmt = select(ConfiguracaoORM).where(ConfiguracaoORM.chave == chave)
            row = uow.session.scalars(stmt).first()
            if row is None:
                return False
            uow.session.delete(row)
            uow.session.flush()
            return True

    def _uow(self) -> UnitOfWork:
        return UnitOfWork(self._session)

    @staticmethod
    def _to_record(row: ConfiguracaoORM) -> ConfiguracaoRecord:
        return ConfiguracaoRecord(
            id=int(row.id),
            chave=row.chave,
            valor=row.valor,
            atualizado_em=row.atualizado_em,
        )
