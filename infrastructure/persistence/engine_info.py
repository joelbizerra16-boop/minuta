"""Metadados do Engine ativo (fonte unica de verdade do banco)."""

from __future__ import annotations

from pathlib import Path

from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session

from infrastructure.persistence.sql_compat import normalize_dialect


def get_engine_dialect(engine: Engine) -> str:
    return normalize_dialect(engine.dialect.name)


def is_sqlite_engine(engine: Engine) -> bool:
    return get_engine_dialect(engine) == "sqlite"


def is_postgresql_engine(engine: Engine) -> bool:
    return get_engine_dialect(engine) == "postgresql"


def get_sqlite_database_path(engine: Engine) -> Path | None:
    """Resolve o caminho do arquivo SQLite a partir do Engine (sem parsing de URL textual)."""
    if not is_sqlite_engine(engine):
        return None
    database = str(engine.url.database or "").strip()
    if not database or database == ":memory:":
        return None
    db_path = Path(database).expanduser()
    if not db_path.is_absolute():
        db_path = db_path.resolve()
    return db_path


def get_dialect_name(session: Session | None = None, *, engine: Engine | None = None) -> str:
    if session is not None and session.bind is not None:
        return get_engine_dialect(session.bind)
    if engine is not None:
        return get_engine_dialect(engine)
    from infrastructure.database import get_engine

    return get_engine_dialect(get_engine())
