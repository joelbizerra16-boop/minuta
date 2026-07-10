from __future__ import annotations

import logging
from pathlib import Path

from sqlalchemy import create_engine, event
from sqlalchemy.engine import Engine, make_url
from sqlalchemy.orm import Session, sessionmaker

from infrastructure.persistence.bootstrap_log import log_engine_configured
from infrastructure.persistence.engine_info import is_sqlite_engine

_LOGGER = logging.getLogger("minuta.persistence.database")

_engine: Engine | None = None
_session_factory: sessionmaker[Session] | None = None
_data_root: Path | None = None
_pdf_storage_dir: Path | None = None
_xml_storage_dir: Path | None = None
_configured_database_url: str | None = None
_database_initialized = False
_engine_create_count = 0

# Adequado ao Streamlit Community Cloud: poucas conexoes concorrentes por processo.
_DEFAULT_POOL_SIZE = 3
_DEFAULT_MAX_OVERFLOW = 5
_DEFAULT_POOL_TIMEOUT = 30
_DEFAULT_POOL_RECYCLE = 300


def _normalize_database_url(database_url: str) -> str:
    return make_url(database_url).render_as_string(hide_password=False)


def get_engine_create_count() -> int:
    return _engine_create_count


def is_database_initialized() -> bool:
    return _database_initialized


def get_configured_database_url() -> str | None:
    return _configured_database_url


def reset_database_state() -> None:
    """Utilitario de teste: descarta engine e estado de inicializacao."""
    global _engine, _session_factory, _data_root, _pdf_storage_dir, _xml_storage_dir
    global _configured_database_url, _database_initialized, _engine_create_count

    if _engine is not None:
        _engine.dispose()
    _engine = None
    _session_factory = None
    _data_root = None
    _pdf_storage_dir = None
    _xml_storage_dir = None
    _configured_database_url = None
    _database_initialized = False
    _engine_create_count = 0


def ensure_database_directories(
    *,
    engine: Engine | None = None,
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
    if engine is not None and is_sqlite_engine(engine):
        from infrastructure.persistence.engine_info import get_sqlite_database_path

        sqlite_path = get_sqlite_database_path(engine)
        if sqlite_path is not None:
            sqlite_path.parent.mkdir(parents=True, exist_ok=True)


def configure_database(
    *,
    database_url: str,
    echo: bool = False,
    data_root: Path | None = None,
    pdf_storage_dir: Path | None = None,
    xml_storage_dir: Path | None = None,
) -> Engine:
    global _engine, _session_factory, _data_root, _pdf_storage_dir, _xml_storage_dir
    global _configured_database_url, _database_initialized, _engine_create_count

    normalized_url = _normalize_database_url(database_url)

    if (
        _database_initialized
        and _engine is not None
        and _configured_database_url == normalized_url
    ):
        if data_root is not None:
            _data_root = data_root
        if pdf_storage_dir is not None:
            _pdf_storage_dir = pdf_storage_dir
        if xml_storage_dir is not None:
            _xml_storage_dir = xml_storage_dir
        ensure_database_directories(
            engine=_engine,
            data_root=_data_root,
            pdf_storage_dir=_pdf_storage_dir,
            xml_storage_dir=_xml_storage_dir,
        )
        _LOGGER.debug(
            "database.configure skipped reutilizando engine database_url=%s creates=%s",
            normalized_url,
            _engine_create_count,
        )
        return _engine

    if _engine is not None:
        _engine.dispose()
        _engine = None
        _session_factory = None

    _data_root = data_root
    _pdf_storage_dir = pdf_storage_dir
    _xml_storage_dir = xml_storage_dir
    _configured_database_url = normalized_url

    engine_kwargs: dict[str, object] = {
        "echo": echo,
        "future": True,
        "pool_pre_ping": True,
    }
    if make_url(database_url).get_backend_name() == "postgresql":
        engine_kwargs.update(
            {
                "pool_size": _DEFAULT_POOL_SIZE,
                "max_overflow": _DEFAULT_MAX_OVERFLOW,
                "pool_timeout": _DEFAULT_POOL_TIMEOUT,
                "pool_recycle": _DEFAULT_POOL_RECYCLE,
            }
        )

    _engine = create_engine(database_url, **engine_kwargs)
    _engine_create_count += 1
    _session_factory = sessionmaker(bind=_engine, autoflush=False, autocommit=False, expire_on_commit=False)
    _database_initialized = True

    ensure_database_directories(
        engine=_engine,
        data_root=data_root,
        pdf_storage_dir=pdf_storage_dir,
        xml_storage_dir=xml_storage_dir,
    )

    if is_sqlite_engine(_engine):
        @event.listens_for(_engine, "connect")
        def _set_sqlite_pragma(dbapi_connection, _connection_record) -> None:  # type: ignore[no-untyped-def]
            cursor = dbapi_connection.cursor()
            cursor.execute("PRAGMA foreign_keys=ON")
            cursor.close()

    log_engine_configured(_engine)
    _LOGGER.info(
        "database.configure engine_criado database_url=%s total_creates=%s pool_size=%s",
        normalized_url,
        _engine_create_count,
        _DEFAULT_POOL_SIZE,
    )
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


def resolve_database_url_from_engine(engine: Engine | None = None) -> str:
    """Retorna a URL do Engine ativo (mascarada apenas para logs externos)."""
    active_engine = engine or get_engine()
    return active_engine.url.render_as_string(hide_password=False)
