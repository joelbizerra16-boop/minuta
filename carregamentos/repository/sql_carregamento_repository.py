from __future__ import annotations

from pathlib import Path

import re

from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from carregamentos.models.carregamento import Carregamento, CarregamentoFiltro
from carregamentos.repository.carregamento_mapper import (
    domain_to_orm,
    item_domain_to_orm,
    orm_to_domain,
)
from carregamentos.repository.carregamento_repository import CarregamentoRepository
from infrastructure.database import get_pdf_storage_dir
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.constants import DOC_TIPO_MINUTA, DOC_TIPO_ROMANEIO
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.unit_of_work import UnitOfWork


class SqlCarregamentoRepository(CarregamentoRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    @property
    def storage_dir(self) -> Path:
        return get_pdf_storage_dir() / "carregamentos"

    def list_all(self) -> list[Carregamento]:
        with self._uow() as uow:
            rows = uow.session.scalars(self._base_stmt()).all()
            return [self._to_domain(uow.session, row) for row in rows]

    def get_by_id(self, carregamento_id: int) -> Carregamento | None:
        with self._uow() as uow:
            row = uow.session.scalars(
                self._base_stmt().where(CarregamentoORM.id == carregamento_id)
            ).first()
            return self._to_domain(uow.session, row) if row else None

    def proximo_numero_carregamento(self) -> str:
        with self._uow() as uow:
            return self._proximo_numero_carregamento_in_session(uow.session)

    @staticmethod
    def _proximo_numero_carregamento_in_session(session: Session) -> str:
        max_id = int(session.scalar(select(func.max(CarregamentoORM.id))) or 0)
        max_seq = 0
        for numero in session.scalars(select(CarregamentoORM.numero_carregamento)).all():
            normalized = str(numero or "").strip()
            if re.fullmatch(r"\d{6}", normalized):
                max_seq = max(max_seq, int(normalized))
        return f"{max(max_id, max_seq) + 1:06d}"

    def get_by_numero(self, numero_carregamento: str) -> Carregamento | None:
        normalized = str(numero_carregamento or "").strip()
        if not normalized:
            return None
        with self._uow() as uow:
            row = uow.session.scalars(
                self._base_stmt().where(CarregamentoORM.numero_carregamento == normalized)
            ).first()
            return self._to_domain(uow.session, row) if row else None

    def save(self, carregamento: Carregamento) -> Carregamento:
        with self._uow() as uow:
            return self._save_in_session(uow.session, carregamento)

    def _save_in_session(self, session: Session, carregamento: Carregamento) -> Carregamento:
        usuario_id = self._resolve_usuario_id(session, carregamento)
        row = session.get(CarregamentoORM, carregamento.id) if carregamento.id > 0 else None
        if row is None:
            row = domain_to_orm(carregamento, usuario_id=usuario_id)
            session.add(row)
            session.flush()
            carregamento.id = int(row.id)
        else:
            domain_to_orm(carregamento, row, usuario_id=usuario_id)

        if carregamento.itens:
            for item in list(row.itens):
                session.delete(item)
            session.flush()
            for index, item in enumerate(carregamento.itens, start=1):
                item_row = item_domain_to_orm(item, int(row.id), index)
                session.add(item_row)

        self._sync_document_paths(session, row, carregamento, usuario_id)
        session.flush()
        reloaded = session.scalars(self._base_stmt().where(CarregamentoORM.id == row.id)).one()
        domain = self._to_domain(session, reloaded)
        if carregamento.itens and not domain.itens:
            from dataclasses import replace

            domain = replace(domain, itens=list(carregamento.itens))
        return domain

    def registrar_impressao(self, session: Session, carregamento_id: int, usuario_id: int) -> Carregamento:
        row = session.get(CarregamentoORM, carregamento_id)
        if row is None:
            raise ValueError("Carregamento nao encontrado.")
        from datetime import datetime, timezone

        row.quantidade_impressoes = int(row.quantidade_impressoes or 0) + 1
        row.ultima_impressao_em = datetime.now(timezone.utc)
        row.ultima_impressao_usuario_id = usuario_id
        session.flush()
        reloaded = session.scalars(self._base_stmt().where(CarregamentoORM.id == row.id)).one()
        return self._to_domain(session, reloaded)

    def sync_document_hashes(
        self,
        session: Session,
        carregamento_id: int,
        hashes: dict[str, str],
    ) -> None:
        for tipo, digest in hashes.items():
            documento = session.scalars(
                select(DocumentoORM).where(
                    DocumentoORM.carregamento_id == carregamento_id,
                    DocumentoORM.tipo == tipo,
                )
            ).first()
            if documento is not None:
                documento.hash_sha256 = digest
        session.flush()

    def search(self, filtro: CarregamentoFiltro) -> list[Carregamento]:
        results = self.list_all()
        if filtro.data_inicial:
            results = [item for item in results if item.data >= filtro.data_inicial]
        if filtro.data_final:
            results = [item for item in results if item.data <= filtro.data_final]
        return sorted(results, key=lambda item: (item.data, item.hora, item.id), reverse=True)

    def _base_stmt(self):
        return (
            select(CarregamentoORM)
            .options(
                selectinload(CarregamentoORM.itens),
                selectinload(CarregamentoORM.documentos),
            )
            .order_by(CarregamentoORM.id)
        )

    def _to_domain(self, session: Session, row: CarregamentoORM | None) -> Carregamento:
        if row is None:
            raise ValueError("Carregamento nao encontrado.")
        usuario = session.get(UsuarioORM, row.usuario_id)
        usuario_login = usuario.usuario if usuario else "sistema"
        ultima_usuario = None
        if row.ultima_impressao_usuario_id:
            ultima_row = session.get(UsuarioORM, row.ultima_impressao_usuario_id)
            ultima_usuario = ultima_row.usuario if ultima_row else None
        return orm_to_domain(
            row,
            list(row.itens),
            list(row.documentos),
            usuario_login=usuario_login,
            ultima_impressao_usuario=ultima_usuario,
        )

    def _resolve_usuario_id(self, session: Session, carregamento: Carregamento) -> int:
        if carregamento.usuario_id:
            return int(carregamento.usuario_id)
        if carregamento.usuario:
            row = session.scalars(
                select(UsuarioORM).where(UsuarioORM.usuario == str(carregamento.usuario).strip().lower())
            ).first()
            if row is not None:
                return int(row.id)
        admin = session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
        if admin is None:
            raise RuntimeError("Nenhum usuario disponivel para vincular o carregamento.")
        return int(admin.id)

    def _sync_document_paths(
        self,
        session: Session,
        row: CarregamentoORM,
        carregamento: Carregamento,
        usuario_id: int,
    ) -> None:
        mapping = {
            DOC_TIPO_MINUTA: carregamento.minuta_pdf_path,
            DOC_TIPO_ROMANEIO: carregamento.romaneio_pdf_path,
        }
        for tipo, relative_path in mapping.items():
            if not relative_path:
                continue
            documento = session.scalars(
                select(DocumentoORM).where(
                    DocumentoORM.carregamento_id == row.id,
                    DocumentoORM.tipo == tipo,
                )
            ).first()
            if documento is None:
                documento = DocumentoORM(
                    carregamento_id=int(row.id),
                    usuario_id=usuario_id,
                    tipo=tipo,
                    caminho_arquivo=relative_path,
                    nome_arquivo=Path(relative_path).name,
                    hash_sha256="0" * 64,
                )
                session.add(documento)
            else:
                documento.caminho_arquivo = relative_path
                documento.nome_arquivo = Path(relative_path).name

    def _uow(self) -> UnitOfWork:
        return UnitOfWork(self._session)
