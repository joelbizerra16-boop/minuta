from __future__ import annotations

from dataclasses import dataclass

from auth.models.usuario import Usuario, UsuarioPublico, utc_now_iso
from auth.repository.usuario_repository import UsuarioRepository
from auth.security.password import verify_password


@dataclass(frozen=True)
class AuthResult:
    success: bool
    user: UsuarioPublico | None = None
    error_message: str = ""
    blocked: bool = False


class AuthService:
    INVALID_CREDENTIALS_MESSAGE = "Usuario ou senha incorretos."
    BLOCKED_MESSAGE = "Seu usuario esta bloqueado. Entre em contato com o administrador."

    def __init__(self, repository: UsuarioRepository):
        self._repository = repository

    def authenticate(self, username: str, password: str) -> AuthResult:
        normalized_username = str(username or "").strip().lower()
        normalized_password = str(password or "")

        if not normalized_username or not normalized_password:
            return AuthResult(success=False, error_message=self.INVALID_CREDENTIALS_MESSAGE)

        usuario = self._repository.get_by_username(normalized_username)
        if usuario is None or not usuario.ativo:
            return AuthResult(success=False, error_message=self.INVALID_CREDENTIALS_MESSAGE)

        if usuario.bloqueado:
            return AuthResult(success=False, error_message=self.BLOCKED_MESSAGE, blocked=True)

        if not verify_password(normalized_password, usuario.senha_hash):
            return AuthResult(success=False, error_message=self.INVALID_CREDENTIALS_MESSAGE)

        usuario.ultimo_login = utc_now_iso()
        saved_user = self._repository.save(usuario)
        return AuthResult(success=True, user=UsuarioPublico.from_usuario(saved_user))
