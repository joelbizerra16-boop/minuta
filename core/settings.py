from __future__ import annotations

import os
from dataclasses import dataclass
from enum import Enum
from functools import lru_cache
from pathlib import Path

from sqlalchemy.engine import make_url

from infrastructure.persistence.sql_compat import normalize_dialect


class StorageBackend(str, Enum):
    SQL = "sql"
    JSON = "json"
    DUAL = "dual"


def _default_database_url() -> str:
    data_root = (Path.cwd() / "data").resolve()
    db_path = data_root / "minuta.db"
    return f"sqlite:///{db_path.as_posix()}"


def derive_storage_paths(database_url: str) -> tuple[Path, Path, Path]:
    """Deriva data_root e diretorios de arquivo a partir de MINUTA_DATABASE_URL."""
    url = make_url(database_url)
    driver = normalize_dialect(url.drivername or "")

    if driver == "sqlite":
        database = str(url.database or "").strip() or ":memory:"
        if database == ":memory:":
            data_root = (Path.cwd() / "data").resolve()
        else:
            data_root = Path(database).expanduser().resolve().parent
    else:
        data_root = (Path.cwd() / "data").resolve()

    return (
        data_root,
        data_root / "documentos",
        data_root / "xml_storage",
    )


@dataclass(frozen=True)
class AppSettings:
    storage_backend: StorageBackend
    database_url: str
    data_root: Path
    pdf_storage_dir: Path
    xml_storage_dir: Path
    echo_sql: bool


@lru_cache(maxsize=1)
def get_settings() -> AppSettings:
    raw_backend = str(os.getenv("MINUTA_STORAGE_BACKEND", StorageBackend.SQL.value) or "").strip().lower()
    try:
        storage_backend = StorageBackend(raw_backend)
    except ValueError:
        storage_backend = StorageBackend.SQL

    database_url = str(os.getenv("MINUTA_DATABASE_URL", _default_database_url()) or _default_database_url()).strip()
    data_root, pdf_storage_dir, xml_storage_dir = derive_storage_paths(database_url)
    echo_sql = str(os.getenv("MINUTA_SQL_ECHO", "0")).strip().lower() in {"1", "true", "yes"}

    return AppSettings(
        storage_backend=storage_backend,
        database_url=database_url,
        data_root=data_root,
        pdf_storage_dir=pdf_storage_dir,
        xml_storage_dir=xml_storage_dir,
        echo_sql=echo_sql,
    )
