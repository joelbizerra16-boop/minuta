"""Utilitarios de mapeamento entre dominio JSON e ORM SQL de usuarios."""

from __future__ import annotations

import uuid
from datetime import datetime, timezone

from auth.models.usuario import PERFIL_ADMIN, PERFIL_OPERADOR, Usuario
from infrastructure.models.usuario import UsuarioORM


def parse_iso_datetime(value: str | datetime | None) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc)
        return value
    raw = str(value).strip()
    if not raw:
        return None
    normalized = raw.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed


def format_iso_datetime(value: datetime | None) -> str | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.replace(microsecond=0).isoformat()


def resolve_perfil_codigo(perfil: str) -> str:
    normalized = str(perfil or PERFIL_OPERADOR).strip().upper()
    if normalized not in {PERFIL_ADMIN, PERFIL_OPERADOR}:
        return PERFIL_OPERADOR
    return normalized


def domain_to_orm(usuario: Usuario, perfil_id: int, row: UsuarioORM | None = None) -> UsuarioORM:
    target = row or UsuarioORM()
    if not target.uuid:
        target.uuid = str(uuid.uuid4())
    if usuario.id > 0:
        target.id = int(usuario.id)
    target.nome = usuario.nome
    target.usuario = usuario.usuario
    target.senha_hash = usuario.senha_hash
    target.perfil_id = perfil_id
    target.perfil = resolve_perfil_codigo(usuario.perfil)
    target.bloqueado = bool(usuario.bloqueado)
    target.ativo = bool(usuario.ativo)
    target.excluido_em = None if usuario.ativo else datetime.now(timezone.utc)
    criado_em = parse_iso_datetime(usuario.criado_em)
    if criado_em is not None:
        target.criado_em = criado_em
    target.ultimo_login = parse_iso_datetime(usuario.ultimo_login)
    return target


def orm_to_domain(row: UsuarioORM) -> Usuario:
    return Usuario(
        id=int(row.id),
        nome=row.nome,
        usuario=row.usuario,
        senha_hash=row.senha_hash,
        perfil=row.perfil,
        ativo=bool(row.ativo),
        bloqueado=bool(row.bloqueado),
        criado_em=format_iso_datetime(row.criado_em) or "",
        ultimo_login=format_iso_datetime(row.ultimo_login),
    )
