from __future__ import annotations

from pathlib import Path

from sqlalchemy import create_engine, event
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

_engine: Engine | None = None
_session_factory: sessionmaker[Session] | None = None
_data_root: Path | None = None
_pdf_storage_dir: Path | None = None
_xml_storage_dir: Path | None = None


def ensure_database_directories(
    database_url: str,
    data_root: Path | None = None,
    pdf_storage_dir: Path | None = None,
    xml_storage_dir: Path | None = None,
) -> None:
    if pdf_storage_dir is not None:
        pdf_storage_dir.mkdir(parents=True, exist_ok=True)
    if xml_storage_dir is not None:
        xml_storage_dir.mkdir(parents=True, exist_ok=True)
    if data_root is not None:
        data_root.mkdir(parents=True, exist_ok=True)
    if database_url.startswith("sqlite:///"):
        sqlite_path = database_url.replace("sqlite:///", "", 1)
        Path(sqlite_path).parent.mkdir(parents=True, exist_ok=True)


def configure_database(
    *,
    database_url: str,
    echo: bool = False,
    data_root: Path | None = None,
    pdf_storage_dir: Path | None = None,
    xml_storage_dir: Path | None = None,
) -> Engine:
    global _engine, _session_factory, _data_root, _pdf_storage_dir, _xml_storage_dir

    _data_root = data_root
    _pdf_storage_dir = pdf_storage_dir
    _xml_storage_dir = xml_storage_dir
    ensure_database_directories(
        database_url,
        data_root=data_root,
        pdf_storage_dir=pdf_storage_dir,
        xml_storage_dir=xml_storage_dir,
    )

    _engine = create_engine(database_url, echo=echo, future=True, pool_pre_ping=True)
    _session_factory = sessionmaker(bind=_engine, autoflush=False, autocommit=False, expire_on_commit=False)

    if database_url.startswith("sqlite"):
        @event.listens_for(_engine, "connect")
        def _set_sqlite_pragma(dbapi_connection, _connection_record) -> None:  # type: ignore[no-untyped-def]
            cursor = dbapi_connection.cursor()
            cursor.execute("PRAGMA foreign_keys=ON")
            cursor.close()

    return _engine


def get_engine() -> Engine:
    if _engine is None:
        raise RuntimeError("Database not configured. Call configure_database first.")
    return _engine


def get_session_factory() -> sessionmaker[Session]:
    if _session_factory is None:
        raise RuntimeError("Database not configured. Call configure_database first.")
    return _session_factory


def get_data_root() -> Path:
    if _data_root is None:
        raise RuntimeError("Data root not configured. Call configure_database first.")
    return _data_root


def get_pdf_storage_dir() -> Path:
    if _pdf_storage_dir is None:
        raise RuntimeError("PDF storage dir not configured. Call configure_database first.")
    return _pdf_storage_dir


def get_xml_storage_dir() -> Path:
    if _xml_storage_dir is None:
        raise RuntimeError("XML storage dir not configured. Call configure_database first.")
    return _xml_storage_dir
