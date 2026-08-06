"""Bootstrap idempotente e seguro para schema SQLite."""

from __future__ import annotations

import logging
import threading

from sqlalchemy import MetaData, inspect, text
from sqlalchemy.engine import Engine

_LOGGER = logging.getLogger("minuta.persistence.sqlite_schema")

_BOOTSTRAP_LOCK = threading.Lock()
_READY_DATABASE_KEY: str | None = None


def reset_sqlite_schema_bootstrap_state() -> None:
    """Utilitario de teste: limpa cache de bootstrap por URL."""
    global _READY_DATABASE_KEY
    with _BOOTSTRAP_LOCK:
        _READY_DATABASE_KEY = None


def _database_key(engine: Engine) -> str:
    return engine.url.render_as_string(hide_password=False)


def apply_sqlite_schema(engine: Engine, metadata: MetaData) -> None:
    """
    Aplica create_all(checkfirst=True) de forma idempotente.

    - Uma unica execucao efetiva por URL de banco no processo (apos sucesso).
    - Lock de processo serializa chamadas concorrentes.
    - BEGIN IMMEDIATE reserva lock de escrita SQLite antes do DDL.
    """
    global _READY_DATABASE_KEY

    database_key = _database_key(engine)
    if _READY_DATABASE_KEY == database_key:
        _LOGGER.debug("sqlite.schema_bootstrap skipped database_key=%s", database_key)
        return

    with _BOOTSTRAP_LOCK:
        if _READY_DATABASE_KEY == database_key:
            _LOGGER.debug("sqlite.schema_bootstrap skipped_after_lock database_key=%s", database_key)
            return

        _LOGGER.info("sqlite.schema_bootstrap start database_key=%s", database_key)
        with engine.connect() as connection:
            connection.execute(text("BEGIN IMMEDIATE"))
            try:
                metadata.create_all(bind=connection, checkfirst=True)
                inspector = inspect(connection)
                if "documento_xml" in inspector.get_table_names():
                    columns = {column["name"] for column in inspector.get_columns("documento_xml")}
                    if "conteudo_xml" not in columns:
                        connection.execute(text("ALTER TABLE documento_xml ADD COLUMN conteudo_xml BLOB"))
            except Exception:
                connection.rollback()
                raise
            else:
                connection.commit()

        _READY_DATABASE_KEY = database_key
        _LOGGER.info("sqlite.schema_bootstrap complete database_key=%s", database_key)
