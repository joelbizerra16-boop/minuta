from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from auth.migration.alembic_runner import run_alembic_upgrade
from auth.migration.usuario_comparator import UsuarioComparisonReport, compare_usuarios, validate_usuario
from auth.repository.perfil_seed import seed_perfis
from auth.repository.sql_usuario_repository import SqlUsuarioRepository
from auth.repository.usuario_mapper import domain_to_orm, orm_to_domain
from auth.repository.usuario_repository import JsonUsuarioRepository
from infrastructure.database import configure_database
from infrastructure.models.usuario import UsuarioORM
from infrastructure.unit_of_work import UnitOfWork


@dataclass
class UsuarioMigrationReport:
    started_at: str
    finished_at: str = ""
    duration_seconds: float = 0.0
    backup_path: str = ""
    migrated_count: int = 0
    existing_sql_count: int = 0
    comparison: UsuarioComparisonReport | None = None
    errors: list[str] = field(default_factory=list)
    success: bool = False

    def to_text(self) -> str:
        lines = [
            "=== RELATORIO MIGRACAO USUARIOS M1 ===",
            f"Inicio: {self.started_at}",
            f"Fim: {self.finished_at}",
            f"Duracao (s): {self.duration_seconds:.3f}",
            f"Backup: {self.backup_path}",
            f"Migrados: {self.migrated_count}",
            f"Existentes SQL antes: {self.existing_sql_count}",
            f"Resultado: {'SUCESSO' if self.success else 'FALHA'}",
        ]
        if self.comparison is not None:
            lines.extend(
                [
                    f"JSON total: {self.comparison.json_total}",
                    f"SQL total: {self.comparison.sql_total}",
                    f"Conferidos: {self.comparison.matched}",
                    f"Divergentes: {self.comparison.divergent}",
                    f"Somente JSON: {', '.join(self.comparison.json_only) or '-'}",
                    f"Somente SQL: {', '.join(self.comparison.sql_only) or '-'}",
                    f"Invalidos: {', '.join(self.comparison.invalid_users) or '-'}",
                    f"Duplicidades: {', '.join(self.comparison.duplicate_usernames) or '-'}",
                ]
            )
            for diff in self.comparison.diffs:
                lines.append(
                    f"  divergencia [{diff.username}] {diff.field}: json={diff.json_value!r} sql={diff.sql_value!r}"
                )
        for error in self.errors:
            lines.append(f"ERRO: {error}")
        return "\n".join(lines)


def backup_usuarios_json(json_path: Path, backup_dir: Path | None = None) -> Path:
    if not json_path.is_file():
        raise FileNotFoundError(f"Arquivo JSON nao encontrado: {json_path}")
    target_dir = backup_dir or json_path.parent / "backups"
    target_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    backup_path = target_dir / f"usuarios.json.bak.{stamp}"
    shutil.copy2(json_path, backup_path)
    return backup_path


def migrate_usuarios_from_json(
    json_path: Path,
    *,
    database_url: str,
    data_root: Path,
    pdf_storage_dir: Path,
    create_backup: bool = True,
) -> UsuarioMigrationReport:
    started = datetime.now(timezone.utc)
    report = UsuarioMigrationReport(started_at=started.isoformat())
    original_bytes: bytes | None = None

    try:
        if create_backup:
            report.backup_path = str(backup_usuarios_json(json_path))
        original_bytes = json_path.read_bytes()

        json_repo = JsonUsuarioRepository(json_path)
        json_users = json_repo.list_all(include_inactive=True)
        if not json_users:
            report.errors.append("Nenhum usuario encontrado no JSON.")
            return report

        for user in json_users:
            validation_errors = validate_usuario(user)
            if validation_errors:
                report.errors.append(f"Usuario invalido {user.usuario}: {', '.join(validation_errors)}")
        if report.errors:
            return report

        configure_database(
            database_url=database_url,
            data_root=data_root,
            pdf_storage_dir=pdf_storage_dir,
        )
        run_alembic_upgrade("head")

        sql_repo = SqlUsuarioRepository()
        report.existing_sql_count = len(sql_repo.list_all(include_inactive=True))

        try:
            with UnitOfWork() as uow:
                perfil_map = seed_perfis(uow.session)
                for usuario in json_users:
                    perfil_id = perfil_map.get(usuario.perfil, perfil_map["OPERADOR"])
                    row = uow.session.get(UsuarioORM, usuario.id)
                    if row is None:
                        row = domain_to_orm(usuario, perfil_id)
                        row.id = usuario.id
                        uow.session.add(row)
                    else:
                        domain_to_orm(usuario, perfil_id, row)
                uow.session.flush()

                sql_users = []
                for user in json_users:
                    row = uow.session.get(UsuarioORM, user.id)
                    if row is not None:
                        sql_users.append(orm_to_domain(row))
                report.comparison = compare_usuarios(json_users, sql_users)
                if not report.comparison.success:
                    report.errors.append("Divergencia detectada entre JSON e SQL. Transacao revertida.")
                    raise RuntimeError("Migracao invalidada por divergencia JSON x SQL.")
        except RuntimeError:
            report.success = False
            return report

        report.migrated_count = len(json_users)
        report.success = True
    except Exception as exc:  # noqa: BLE001
        report.errors.append(str(exc))
        report.success = False
    finally:
        if original_bytes is not None and json_path.read_bytes() != original_bytes:
            report.errors.append("JSON original foi alterado e sera restaurado.")
            json_path.write_bytes(original_bytes)
        finished = datetime.now(timezone.utc)
        report.finished_at = finished.isoformat()
        report.duration_seconds = (finished - started).total_seconds()

    return report
