from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from sqlalchemy import inspect

from auth.repository.perfil_seed import seed_perfis
from infrastructure.database import get_engine
from infrastructure.models.perfil import PerfilORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.unit_of_work import UnitOfWork

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def run_alembic_upgrade(revision: str = "head") -> None:
    """Garante schema M1 no banco configurado via configure_database()."""
    if revision != "head":
        raise ValueError(f"Revisao nao suportada: {revision}")
    ensure_m1_auth_schema()


def run_alembic_downgrade(revision: str = "base") -> None:
    if revision != "base":
        raise ValueError(f"Revisao nao suportada: {revision}")
    drop_m1_auth_schema()


def run_alembic_cli_upgrade(revision: str = "head") -> None:
    """Executa Alembic CLI (requer MINUTA_DATABASE_URL no ambiente)."""
    alembic_bin = shutil.which("alembic")
    if not alembic_bin:
        raise RuntimeError("CLI alembic nao encontrada no PATH.")
    result = subprocess.run(
        [alembic_bin, "upgrade", revision],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"Alembic upgrade {revision} falhou: {detail}")


def ensure_m1_auth_schema() -> None:
    engine = get_engine()
    PerfilORM.__table__.create(engine, checkfirst=True)
    UsuarioORM.__table__.create(engine, checkfirst=True)
    with UnitOfWork() as uow:
        seed_perfis(uow.session)


def drop_m1_auth_schema() -> None:
    engine = get_engine()
    inspector = inspect(engine)
    if inspector.has_table("usuario"):
        UsuarioORM.__table__.drop(engine, checkfirst=True)
    if inspector.has_table("perfil"):
        PerfilORM.__table__.drop(engine, checkfirst=True)
