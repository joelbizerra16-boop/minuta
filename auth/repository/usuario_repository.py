from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from auth.models.usuario import Usuario


class UsuarioRepository(ABC):
    @abstractmethod
    def list_all(self, include_inactive: bool = False) -> list[Usuario]:
        raise NotImplementedError

    @abstractmethod
    def get_by_id(self, user_id: int) -> Usuario | None:
        raise NotImplementedError

    @abstractmethod
    def get_by_username(self, username: str) -> Usuario | None:
        raise NotImplementedError

    @abstractmethod
    def save(self, usuario: Usuario) -> Usuario:
        raise NotImplementedError

    @abstractmethod
    def delete_logical(self, user_id: int) -> Usuario | None:
        raise NotImplementedError

    @abstractmethod
    def ensure_default_admin(self, username: str, password: str, nome: str) -> None:
        raise NotImplementedError


class JsonUsuarioRepository(UsuarioRepository):
    def __init__(self, json_path: Path):
        self._json_path = json_path

    def list_all(self, include_inactive: bool = False) -> list[Usuario]:
        usuarios = self._load_usuarios()
        if include_inactive:
            return usuarios
        return [usuario for usuario in usuarios if usuario.ativo]

    def get_by_id(self, user_id: int) -> Usuario | None:
        for usuario in self._load_usuarios():
            if usuario.id == user_id:
                return usuario
        return None

    def get_by_username(self, username: str) -> Usuario | None:
        normalized = str(username or "").strip().lower()
        if not normalized:
            return None
        for usuario in self._load_usuarios(include_inactive=True):
            if usuario.usuario == normalized:
                return usuario
        return None

    def save(self, usuario: Usuario) -> Usuario:
        usuarios = self._load_usuarios(include_inactive=True)
        updated = False
        for index, current in enumerate(usuarios):
            if current.id == usuario.id:
                usuarios[index] = usuario
                updated = True
                break
        if not updated:
            if usuario.id <= 0:
                usuario = Usuario(
                    id=self._next_id(),
                    nome=usuario.nome,
                    usuario=usuario.usuario,
                    senha_hash=usuario.senha_hash,
                    perfil=usuario.perfil,
                    ativo=usuario.ativo,
                    bloqueado=usuario.bloqueado,
                    criado_em=usuario.criado_em,
                    ultimo_login=usuario.ultimo_login,
                )
            usuarios.append(usuario)
        self._write_usuarios(usuarios)
        return usuario

    def delete_logical(self, user_id: int) -> Usuario | None:
        usuario = self.get_by_id(user_id)
        if usuario is None:
            return None
        usuario.ativo = False
        return self.save(usuario)

    def ensure_default_admin(self, username: str, password: str, nome: str) -> None:
        from auth.security.password import hash_password

        usuarios = self._load_usuarios(include_inactive=True)
        if usuarios:
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
        self._write_usuarios([admin])

    def _load_usuarios(self, include_inactive: bool = True) -> list[Usuario]:
        import json

        self._json_path.parent.mkdir(parents=True, exist_ok=True)
        if not self._json_path.is_file():
            return []

        try:
            payload = json.loads(self._json_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            return []

        raw_users = payload.get("usuarios", []) if isinstance(payload, dict) else []
        usuarios = [Usuario.from_dict(item) for item in raw_users if isinstance(item, dict)]
        if include_inactive:
            return sorted(usuarios, key=lambda item: item.id)
        return sorted([usuario for usuario in usuarios if usuario.ativo], key=lambda item: item.id)

    def _write_usuarios(self, usuarios: list[Usuario]) -> None:
        import json

        self._json_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"usuarios": [usuario.to_dict() for usuario in sorted(usuarios, key=lambda item: item.id)]}
        self._json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _next_id(self) -> int:
        usuarios = self._load_usuarios(include_inactive=True)
        if not usuarios:
            return 1
        return max(usuario.id for usuario in usuarios) + 1

    def create(self, usuario: Usuario) -> Usuario:
        usuarios = self._load_usuarios(include_inactive=True)
        usuario.id = self._next_id()
        usuarios.append(usuario)
        self._write_usuarios(usuarios)
        return usuario
