from __future__ import annotations

from pathlib import Path

from auth.migration.usuario_comparator import compare_usuarios
from auth.models.usuario import Usuario
from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.repository.usuario_repository import JsonUsuarioRepository


class DualWriteValidationError(RuntimeError):
    pass


class DualUsuarioRepository(JsonUsuarioRepository):
    """
    Fonte oficial: JSON.
    Persistencia paralela: SQL com comparacao automatica.
    """

    def __init__(self, json_path: Path, sql_repository: SqlUsuarioRepository, *, strict_validation: bool = True):
        super().__init__(json_path)
        self._sql_repository = sql_repository
        self._strict_validation = strict_validation

    def save(self, usuario: Usuario) -> Usuario:
        saved = super().save(usuario)
        self._mirror_and_validate(saved)
        return saved

    def delete_logical(self, user_id: int) -> Usuario | None:
        deleted = super().delete_logical(user_id)
        if deleted is not None:
            self._mirror_and_validate(deleted)
        return deleted

    def ensure_default_admin(self, username: str, password: str, nome: str) -> None:
        super().ensure_default_admin(username, password, nome)
        self.sync_all_from_json()

    def create(self, usuario: Usuario) -> Usuario:
        created = super().create(usuario)
        self._mirror_and_validate(created)
        return created

    def sync_all_from_json(self) -> None:
        json_users = self.list_all(include_inactive=True)
        for usuario in json_users:
            self._sql_repository.save(usuario)
        self._validate_all(json_users)

    def _mirror_and_validate(self, usuario: Usuario) -> None:
        sql_saved = self._sql_repository.save(usuario)
        sql_user = self._sql_repository.get_by_username(sql_saved.usuario)
        if sql_user is None:
            message = f"Usuario {usuario.usuario} nao encontrado no SQL apos espelhamento."
            if self._strict_validation:
                raise DualWriteValidationError(message)
            return
        report = compare_usuarios([usuario], [sql_user])
        if not report.success:
            details = ", ".join(f"{diff.field}" for diff in report.diffs)
            message = f"Divergencia JSON x SQL para {usuario.usuario}: {details}"
            if self._strict_validation:
                raise DualWriteValidationError(message)

    def _validate_all(self, json_users: list[Usuario]) -> None:
        sql_users = self._sql_repository.list_all(include_inactive=True)
        report = compare_usuarios(json_users, sql_users)
        if not report.success and self._strict_validation:
            raise DualWriteValidationError(
                f"Divergencia global JSON x SQL: divergentes={report.divergent}, "
                f"json_only={report.json_only}, sql_only={report.sql_only}"
            )
