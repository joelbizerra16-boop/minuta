from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

PERFIL_ADMIN = "ADMIN"
PERFIL_OPERADOR = "OPERADOR"
PERFIS_VALIDOS = {PERFIL_ADMIN, PERFIL_OPERADOR}


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


@dataclass
class Usuario:
    id: int
    nome: str
    usuario: str
    senha_hash: str
    perfil: str = PERFIL_OPERADOR
    ativo: bool = True
    bloqueado: bool = False
    criado_em: str = field(default_factory=utc_now_iso)
    ultimo_login: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "nome": self.nome,
            "usuario": self.usuario,
            "senha_hash": self.senha_hash,
            "perfil": self.perfil,
            "ativo": self.ativo,
            "bloqueado": self.bloqueado,
            "criado_em": self.criado_em,
            "ultimo_login": self.ultimo_login,
        }

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> Usuario:
        return cls(
            id=int(payload.get("id", 0)),
            nome=str(payload.get("nome", "")).strip(),
            usuario=str(payload.get("usuario", "")).strip().lower(),
            senha_hash=str(payload.get("senha_hash", "")),
            perfil=str(payload.get("perfil", PERFIL_OPERADOR)).strip().upper(),
            ativo=bool(payload.get("ativo", True)),
            bloqueado=bool(payload.get("bloqueado", False)),
            criado_em=str(payload.get("criado_em", utc_now_iso())),
            ultimo_login=payload.get("ultimo_login"),
        )


@dataclass(frozen=True)
class UsuarioPublico:
    id: int
    nome: str
    usuario: str
    perfil: str
    ativo: bool
    bloqueado: bool
    criado_em: str
    ultimo_login: str | None

    @classmethod
    def from_usuario(cls, usuario: Usuario) -> UsuarioPublico:
        return cls(
            id=usuario.id,
            nome=usuario.nome,
            usuario=usuario.usuario,
            perfil=usuario.perfil,
            ativo=usuario.ativo,
            bloqueado=usuario.bloqueado,
            criado_em=usuario.criado_em,
            ultimo_login=usuario.ultimo_login,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "nome": self.nome,
            "usuario": self.usuario,
            "perfil": self.perfil,
            "ativo": self.ativo,
            "bloqueado": self.bloqueado,
            "criado_em": self.criado_em,
            "ultimo_login": self.ultimo_login,
        }
