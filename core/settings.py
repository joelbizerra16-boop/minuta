from __future__ import annotations

import os
from dataclasses import dataclass
from enum import Enum
from functools import lru_cache
from pathlib import Path


class StorageBackend(str, Enum):
    SQL = "sql"
    JSON = "json"
    DUAL = "dual"


def _default_database_url() -> str:
    default_sqlite = Path(os.getenv("MINUTA_DATA_ROOT", r"C:\MinutaData")) / "minuta_dev.db"
    return f"sqlite:///{default_sqlite.as_posix()}"


@dataclass(frozen=True)
class AppSettings:
    storage_backend: StorageBackend
    database_url: str
    data_root: Path
    pdf_storage_dir: Path
    echo_sql: bool


@lru_cache(maxsize=1)
def get_settings() -> AppSettings:
    raw_backend = str(os.getenv("MINUTA_STORAGE_BACKEND", StorageBackend.SQL.value) or "").strip().lower()
    try:
        storage_backend = StorageBackend(raw_backend)
    except ValueError:
        storage_backend = StorageBackend.SQL

    data_root = Path(os.getenv("MINUTA_DATA_ROOT", r"C:\MinutaData"))
    pdf_storage_dir = Path(os.getenv("MINUTA_PDF_STORAGE_DIR", str(data_root / "documentos")))
    database_url = str(os.getenv("MINUTA_DATABASE_URL", _default_database_url()) or _default_database_url()).strip()
    echo_sql = str(os.getenv("MINUTA_SQL_ECHO", "0")).strip().lower() in {"1", "true", "yes"}

    return AppSettings(
        storage_backend=storage_backend,
        database_url=database_url,
        data_root=data_root,
        pdf_storage_dir=pdf_storage_dir,
        echo_sql=echo_sql,
    )
