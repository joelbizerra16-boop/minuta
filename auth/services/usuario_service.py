from __future__ import annotations

import re

from auth.models.usuario import PERFIL_ADMIN, PERFIL_OPERADOR, PERFIS_VALIDOS, Usuario, UsuarioPublico, utc_now_iso
from auth.repository.usuario_repository import UsuarioRepository
from auth.security.password import hash_password

USERNAME_PATTERN = re.compile(r"^[a-z0-9._-]{3,32}$")


class UsuarioService:
    MIN_PASSWORD_LENGTH = 6

    def __init__(self, repository: UsuarioRepository):
        self._repository = repository

    def list_users(self, include_inactive: bool = True) -> list[UsuarioPublico]:
        return [UsuarioPublico.from_usuario(usuario) for usuario in self._repository.list_all(include_inactive=include_inactive)]

    def get_user(self, user_id: int) -> UsuarioPublico | None:
        usuario = self._repository.get_by_id(user_id)
        if usuario is None:
            return None
        return UsuarioPublico.from_usuario(usuario)

    def create_user(self, nome: str, usuario: str, senha: str, perfil: str) -> UsuarioPublico:
        normalized_name = self._normalize_name(nome)
        normalized_username = self._normalize_username(usuario)
        normalized_profile = self._normalize_profile(perfil)
        self._validate_password(senha)

        if self._repository.get_by_username(normalized_username) is not None:
            raise ValueError("Usuario ja cadastrado.")

        novo_usuario = Usuario(
            id=0,
            nome=normalized_name,
            usuario=normalized_username,
            senha_hash=hash_password(senha),
            perfil=normalized_profile,
            ativo=True,
            bloqueado=False,
            criado_em=utc_now_iso(),
            ultimo_login=None,
        )
        saved = self._repository.save(novo_usuario)
        return UsuarioPublico.from_usuario(saved)

    def update_user(self, user_id: int, nome: str, perfil: str) -> UsuarioPublico:
        usuario = self._require_active_user(user_id)
        usuario.nome = self._normalize_name(nome)
        usuario.perfil = self._normalize_profile(perfil)
        saved = self._repository.save(usuario)
        return UsuarioPublico.from_usuario(saved)

    def change_password(self, user_id: int, new_password: str) -> UsuarioPublico:
        usuario = self._require_active_user(user_id)
        self._validate_password(new_password)
        usuario.senha_hash = hash_password(new_password)
        saved = self._repository.save(usuario)
        return UsuarioPublico.from_usuario(saved)

    def block_user(self, user_id: int) -> UsuarioPublico:
        usuario = self._require_active_user(user_id)
        usuario.bloqueado = True
        saved = self._repository.save(usuario)
        return UsuarioPublico.from_usuario(saved)

    def unblock_user(self, user_id: int) -> UsuarioPublico:
        usuario = self._require_active_user(user_id)
        usuario.bloqueado = False
        saved = self._repository.save(usuario)
        return UsuarioPublico.from_usuario(saved)

    def delete_user(self, user_id: int) -> UsuarioPublico:
        usuario = self._require_active_user(user_id)
        deleted = self._repository.delete_logical(user_id)
        if deleted is None:
            raise ValueError("Usuario nao encontrado.")
        return UsuarioPublico.from_usuario(deleted)

    def _require_active_user(self, user_id: int) -> Usuario:
        usuario = self._repository.get_by_id(user_id)
        if usuario is None or not usuario.ativo:
            raise ValueError("Usuario nao encontrado.")
        return usuario

    def _normalize_name(self, nome: str) -> str:
        normalized = str(nome or "").strip()
        if len(normalized) < 3:
            raise ValueError("Informe o nome completo do usuario.")
        return normalized

    def _normalize_username(self, usuario: str) -> str:
        normalized = str(usuario or "").strip().lower()
        if not USERNAME_PATTERN.fullmatch(normalized):
            raise ValueError("Usuario invalido. Use de 3 a 32 caracteres (letras, numeros, ., _ ou -).")
        return normalized

    def _normalize_profile(self, perfil: str) -> str:
        normalized = str(perfil or PERFIL_OPERADOR).strip().upper()
        if normalized not in PERFIS_VALIDOS:
            raise ValueError("Perfil invalido.")
        return normalized

    def _validate_password(self, senha: str) -> None:
        if len(str(senha or "")) < self.MIN_PASSWORD_LENGTH:
            raise ValueError(f"A senha deve ter pelo menos {self.MIN_PASSWORD_LENGTH} caracteres.")
