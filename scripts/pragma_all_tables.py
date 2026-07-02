"""PRAGMA table_info(id) para todas as tabelas do dominio."""
from __future__ import annotations

import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.settings import get_settings

TABLES = (
    "perfil",
    "usuario",
    "motorista",
    "veiculo",
    "destinatario",
    "rota",
    "nota_fiscal",
    "item_nota_fiscal",
    "configuracao",
    "carregamento",
    "item_carregamento",
    "documento",
    "historico_operacional",
    "evento_auditoria",
)


def main() -> None:
    path = get_settings().database_url.replace("sqlite:///", "")
    conn = sqlite3.connect(path)
    cur = conn.cursor()
    print(f"database: {Path(path).resolve()}\n")
    for table in TABLES:
        cur.execute(f"PRAGMA table_info([{table}])")
        id_col = next((row for row in cur.fetchall() if row[1] == "id"), None)
        if id_col:
            print(f"{table:24} type={id_col[2]:8} pk={id_col[5]}")
        else:
            print(f"{table:24} id AUSENTE")
    cur.execute("PRAGMA foreign_key_check")
    fk = cur.fetchall()
    print(f"\nforeign_key_check: {len(fk)} violacoes")
    cur.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%'")
    print(f"indices: {cur.fetchone()[0]}")
    cur.execute("SELECT version_num FROM alembic_version")
    print(f"alembic_version: {cur.fetchone()[0]}")
    conn.close()


if __name__ == "__main__":
    main()
