from __future__ import annotations

from sqlalchemy import select
from sqlalchemy.orm import Session

from auth.models.usuario import PERFIL_ADMIN, PERFIL_OPERADOR
from infrastructure.models.perfil import PerfilORM

PERFIS_SEED = (
    (1, PERFIL_ADMIN, "Administrador"),
    (2, PERFIL_OPERADOR, "Operador"),
)


def seed_perfis(session: Session) -> dict[str, int]:
    """Garante perfis ADMIN e OPERADOR sem duplicar."""
    mapping: dict[str, int] = {}
    for perfil_id, codigo, nome in PERFIS_SEED:
        row = session.get(PerfilORM, perfil_id)
        if row is None:
            row = session.scalars(select(PerfilORM).where(PerfilORM.codigo == codigo)).first()
        if row is None:
            row = PerfilORM(id=perfil_id, codigo=codigo, nome=nome)
            session.add(row)
            session.flush()
        mapping[codigo] = int(row.id)
    return mapping


def get_perfil_id(session: Session, perfil_codigo: str) -> int:
    mapping = seed_perfis(session)
    normalized = str(perfil_codigo or PERFIL_OPERADOR).strip().upper()
    if normalized not in mapping:
        normalized = PERFIL_OPERADOR
    return mapping[normalized]
