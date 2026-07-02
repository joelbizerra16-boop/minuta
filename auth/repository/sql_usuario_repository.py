from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from auth.models.usuario import Usuario
from auth.repository.perfil_seed import get_perfil_id, seed_perfis
from auth.repository.usuario_mapper import domain_to_orm, orm_to_domain
from auth.repository.usuario_repository import UsuarioRepository
from auth.security.password import hash_password
from infrastructure.models.usuario import UsuarioORM
from infrastructure.unit_of_work import UnitOfWork


class SqlUsuarioRepository(UsuarioRepository):
    def __init__(self, session: Session | None = None) -> None:
        self._session = session

    def list_all(self, include_inactive: bool = False) -> list[Usuario]:
        with self._uow() as uow:
            stmt = select(UsuarioORM).order_by(UsuarioORM.id)
            rows = uow.session.scalars(stmt).all()
            usuarios = [orm_to_domain(row) for row in rows]
            if include_inactive:
                return usuarios
            return [usuario for usuario in usuarios if usuario.ativo]

    def get_by_id(self, user_id: int) -> Usuario | None:
        with self._uow() as uow:
            row = uow.session.get(UsuarioORM, user_id)
            return orm_to_domain(row) if row else None

    def get_by_username(self, username: str) -> Usuario | None:
        normalized = str(username or "").strip().lower()
        if not normalized:
            return None
        with self._uow() as uow:
            stmt = select(UsuarioORM).where(UsuarioORM.usuario == normalized)
            row = uow.session.scalars(stmt).first()
            return orm_to_domain(row) if row else None

    def save(self, usuario: Usuario) -> Usuario:
        with self._uow() as uow:
            seed_perfis(uow.session)
            row = None
            if usuario.id > 0:
                row = uow.session.get(UsuarioORM, usuario.id)
            perfil_id = get_perfil_id(uow.session, usuario.perfil)
            if row is None:
                row = domain_to_orm(usuario, perfil_id)
                uow.session.add(row)
            else:
                domain_to_orm(usuario, perfil_id, row)
            uow.session.flush()
            return orm_to_domain(row)

    def delete_logical(self, user_id: int) -> Usuario | None:
        usuario = self.get_by_id(user_id)
        if usuario is None:
            return None
        usuario.ativo = False
        return self.save(usuario)

    def ensure_default_admin(self, username: str, password: str, nome: str) -> None:
        with self._uow() as uow:
            seed_perfis(uow.session)
            existing = uow.session.scalars(select(UsuarioORM).limit(1)).first()
            if existing is not None:
                return
            admin = Usuario(
                id=1,
                nome=nome,
                usuario=username.strip().lower(),
                senha_hash=hash_password(password),
                perfil="ADMIN",
                ativo=True,
                bloqueado=False,
            )
            row = domain_to_orm(admin, get_perfil_id(uow.session, "ADMIN"))
            row.id = 1
            uow.session.add(row)
            uow.session.flush()

    def _uow(self) -> UnitOfWork:
        return UnitOfWork(self._session)
