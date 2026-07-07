"""Auditoria read-only: ORM runtime vs schema fisico SQLite."""
from __future__ import annotations

import sqlite3
from pathlib import Path

from sqlalchemy import inspect
from sqlalchemy.schema import CreateTable

from core.bootstrap import configure_application_storage
from core.settings import get_settings
from infrastructure.database import get_engine
from infrastructure.models.configuracao import ConfiguracaoORM
from infrastructure.persistence.engine_info import get_sqlite_database_path


def main() -> None:
    settings = get_settings()
    print("=== RUNTIME SETTINGS ===")
    print(f"database_url: {settings.database_url}")
    print(f"data_root: {settings.data_root}")
    print(f"pdf_storage_dir: {settings.pdf_storage_dir}")

    configure_application_storage()

    engine = get_engine()
    resolved = get_sqlite_database_path(engine)
    print(f"\nengine.url: {engine.url.render_as_string(hide_password=True)}")
    print(f"sqlite_file_resolved: {resolved}")
    if resolved is None:
        print("file_exists: n/a (nao e SQLite em arquivo)")
        return
    print(f"file_exists: {resolved.exists()}")
    if resolved.exists():
        stat = resolved.stat()
        print(f"file_size_bytes: {stat.st_size}")
        print(f"file_mtime: {stat.st_mtime}")

    orm_ddl = str(CreateTable(ConfiguracaoORM.__table__).compile(dialect=engine.dialect))
    print("\n=== ORM DDL (SQLAlchemy compile, SQLite) ===")
    print(orm_ddl)

    inspector = inspect(engine)
    print("\n=== SQLAlchemy inspector: configuracao ===")
    if inspector.has_table("configuracao"):
        for column in inspector.get_columns("configuracao"):
            print(column)
    else:
        print("TABELA AUSENTE")

    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    print("\n=== PRAGMA table_info(configuracao) ===")
    cur.execute("PRAGMA table_info(configuracao)")
    pragma_rows = [dict(row) for row in cur.fetchall()]
    for row in pragma_rows:
        print(row)

    print("\n=== sqlite_master DDL ===")
    cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='configuracao'")
    ddl_row = cur.fetchone()
    print(ddl_row[0] if ddl_row else "TABLE NOT FOUND")

    print("\n=== alembic_version ===")
    try:
        cur.execute("SELECT version_num FROM alembic_version")
        print(cur.fetchall())
    except sqlite3.Error as exc:
        print(f"erro: {exc}")

    print("\n=== sqlite_sequence ===")
    try:
        cur.execute("SELECT name, seq FROM sqlite_sequence ORDER BY name")
        for row in cur.fetchall():
            print(tuple(row))
    except sqlite3.Error as exc:
        print(f"erro: {exc}")

    print("\n=== id column physical summary ===")
    id_col = next((row for row in pragma_rows if row.get("name") == "id"), None)
    print(id_col)

    print("\n=== row counts ===")
    for table in (
        "perfil",
        "usuario",
        "configuracao",
        "carregamento",
        "nota_fiscal",
        "documento",
        "evento_auditoria",
    ):
        try:
            cur.execute(f"SELECT COUNT(*) FROM {table}")
            print(f"{table}: {cur.fetchone()[0]}")
        except sqlite3.Error as exc:
            print(f"{table}: {exc}")

    print("\n=== PK DDL snippets (perfil, usuario, configuracao) ===")
    cur.execute(
        "SELECT name, sql FROM sqlite_master WHERE type='table' AND name IN ('perfil','usuario','configuracao') ORDER BY name"
    )
    for name, sql in cur.fetchall():
        print(f"--- {name} ---")
        print(sql)

    conn.close()


if __name__ == "__main__":
    main()
