"""
Reconstrucao controlada do banco SQLite legado (BigInteger -> Integer PK).

Fluxo:
1. Backup automatico do banco atual
2. Novo banco via create_all() com ORM atual
3. Migracao de dados preservando IDs e FKs
4. Validacao de schema, constraints e autoincrement
5. Substituicao atomica do arquivo de banco
6. Stamp alembic_version = head
"""

from __future__ import annotations

import shutil
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.schema import CreateTable

# Garante import do pacote raiz
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.settings import get_settings
from infrastructure.models import Base

ALEMBIC_HEAD = "m3_0003_integer_surrogate_keys"

# Ordem respeitando dependencias de FK
MIGRATION_TABLES = (
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

TABLES_WITH_INTEGER_PK = (
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


def resolve_sqlite_path(database_url: str) -> Path:
    if not database_url.startswith("sqlite:///"):
        raise RuntimeError(f"Este script suporta apenas SQLite. URL atual: {database_url}")
    return Path(database_url.replace("sqlite:///", "", 1)).resolve()


def backup_database(source: Path, backups_dir: Path) -> Path:
    backups_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    backup_path = backups_dir / f"{source.stem}.bak-{stamp}{source.suffix}"
    shutil.copy2(source, backup_path)
    return backup_path


def list_user_tables(conn: sqlite3.Connection) -> list[str]:
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
    return [row[0] for row in cur.fetchall() if row[0] != "alembic_version"]


def table_columns(conn: sqlite3.Connection, table: str) -> list[str]:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info([{table}])")
    return [row[1] for row in cur.fetchall()]


def count_rows(conn: sqlite3.Connection, table: str) -> int:
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM [{table}]")
    return int(cur.fetchone()[0])


def copy_table_data(source: sqlite3.Connection, target: sqlite3.Connection, table: str) -> int:
    if table not in list_user_tables(source):
        return 0
    if count_rows(source, table) == 0:
        return 0

    src_cols = table_columns(source, table)
    tgt_cols = table_columns(target, table)
    common = [col for col in src_cols if col in tgt_cols]
    if not common:
        return 0

    quoted = ", ".join(f"[{col}]" for col in common)
    placeholders = ", ".join("?" for _ in common)
    select_sql = f"SELECT {quoted} FROM [{table}]"
    insert_sql = f"INSERT INTO [{table}] ({quoted}) VALUES ({placeholders})"

    src_cur = source.cursor()
    tgt_cur = target.cursor()
    rows = src_cur.execute(select_sql).fetchall()
    tgt_cur.executemany(insert_sql, rows)
    target.commit()
    return len(rows)


def create_schema(database_path: Path) -> None:
    if database_path.exists():
        database_path.unlink()
    database_path.parent.mkdir(parents=True, exist_ok=True)
    engine = create_engine(f"sqlite:///{database_path.as_posix()}", future=True)
    Base.metadata.create_all(engine)
    engine.dispose()


def sync_sqlite_sequences(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='sqlite_sequence'")
    if cur.fetchone() is None:
        return
    for table in TABLES_WITH_INTEGER_PK:
        try:
            max_id = cur.execute(f"SELECT COALESCE(MAX(id), 0) FROM [{table}]").fetchone()[0]
        except sqlite3.Error:
            continue
        if int(max_id) <= 0:
            continue
        cur.execute("SELECT 1 FROM sqlite_sequence WHERE name = ?", (table,))
        if cur.fetchone():
            cur.execute("UPDATE sqlite_sequence SET seq = ? WHERE name = ?", (int(max_id), table))
        else:
            cur.execute("INSERT INTO sqlite_sequence (name, seq) VALUES (?, ?)", (table, int(max_id)))
    conn.commit()


def stamp_alembic(conn: sqlite3.Connection, revision: str) -> None:
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS alembic_version (version_num VARCHAR(32) NOT NULL)")
    cur.execute("DELETE FROM alembic_version")
    cur.execute("INSERT INTO alembic_version (version_num) VALUES (?)", (revision,))
    conn.commit()


def validate_integer_primary_keys(conn: sqlite3.Connection) -> list[str]:
    errors: list[str] = []
    cur = conn.cursor()
    for table in TABLES_WITH_INTEGER_PK:
        cur.execute(f"PRAGMA table_info([{table}])")
        id_col = next((row for row in cur.fetchall() if row[1] == "id"), None)
        if id_col is None:
            errors.append(f"{table}: coluna id ausente")
            continue
        col_type = str(id_col[2]).upper()
        is_pk = int(id_col[5]) == 1
        if "INT" not in col_type:
            errors.append(f"{table}: id tipo fisico {col_type} (esperado INTEGER)")
        if not is_pk:
            errors.append(f"{table}: id nao e primary key")
    return errors


def validate_foreign_keys(conn: sqlite3.Connection) -> list[str]:
    conn.execute("PRAGMA foreign_keys=ON")
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_key_check")
    violations = cur.fetchall()
    return [f"FK violation: {row}" for row in violations]


def validate_indexes_and_constraints(conn: sqlite3.Connection) -> dict[str, int]:
    cur = conn.cursor()
    cur.execute("SELECT type, COUNT(*) FROM sqlite_master WHERE type IN ('index','table') GROUP BY type")
    counts = {row[0]: int(row[1]) for row in cur.fetchall()}
    cur.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND sql IS NOT NULL")
    counts["explicit_indexes"] = int(cur.fetchone()[0])
    return counts


def validate_autoincrement_inserts(database_path: Path) -> dict[str, int]:
    """Insere registros de teste sem ID e remove apos validacao."""
    from decimal import Decimal
    import infrastructure.database as db_module
    from auth.repository.sql_usuario_repository import SqlUsuarioRepository
    from carregamentos.models.carregamento import MODALIDADE_VEICULO, STATUS_FINALIZADO, Carregamento
    from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
    from infrastructure.database import configure_database, get_engine
    from infrastructure.repositories.configuracao_repository import ConfiguracaoRecord
    from infrastructure.repositories.documento_repository import DocumentoRecord
    from infrastructure.repositories.historico_repository import HistoricoRecord
    from infrastructure.repositories.nota_fiscal_repository import NotaFiscalRecord
    from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
    from infrastructure.repositories.sql.documento_repository import SqlDocumentoRepository
    from infrastructure.repositories.sql.historico_repository import SqlHistoricoRepository
    from infrastructure.repositories.sql.nota_fiscal_repository import SqlNotaFiscalRepository
    from infrastructure.models.constants import DOC_TIPO_MINUTA
    from infrastructure.unit_of_work import UnitOfWork

    data_root = database_path.parent
    configure_database(
        database_url=f"sqlite:///{database_path.as_posix()}",
        data_root=data_root,
        pdf_storage_dir=data_root / "documentos",
    )

    generated: dict[str, int] = {}

    cfg = SqlConfiguracaoRepository().save(
        ConfiguracaoRecord(id=0, chave="__rebuild_validation__", valor="{}")
    )
    generated["configuracao"] = cfg.id
    assert cfg.id > 0

    admin_id = SqlUsuarioRepository().get_by_username("admin")
    if admin_id is None:
        raise RuntimeError("Usuario admin ausente apos rebuild.")
    usuario_id = int(admin_id.id)

    with UnitOfWork() as uow:
        nf_repo = SqlNotaFiscalRepository(uow.session)
        nf = nf_repo.save(
            NotaFiscalRecord(
                id=0,
                chave_nfe="0" * 44,
                numero_nf="REBUILD-TEST",
                destinatario="Teste Rebuild",
                status_nf="PENDENTE",
                valor_total=Decimal("0"),
                peso_total=Decimal("0"),
                volume_total=Decimal("0"),
            )
        )
        generated["nota_fiscal"] = nf.id
        assert nf.id > 0

        carreg_repo = SqlCarregamentoRepository(uow.session)
        carreg = carreg_repo.save(
            Carregamento(
                id=0,
                numero_carregamento="REBUILD-TEST-001",
                data="2026-06-30",
                hora="12:00:00",
                usuario="admin",
                usuario_id=None,
                motorista="Teste",
                placa="ABC1D23",
                filial="",
                data_saida="--",
                quantidade_nf=0,
                quantidade_itens=0,
                peso_total=0.0,
                status=STATUS_FINALIZADO,
                modalidade=MODALIDADE_VEICULO,
                reentrega=False,
                minuta_pdf_path=None,
                romaneio_pdf_path=None,
                itens=[],
                criado_em="2026-06-30T12:00:00+00:00",
            )
        )
        generated["carregamento"] = carreg.id
        assert carreg.id > 0

        doc_repo = SqlDocumentoRepository(uow.session)
        doc = doc_repo.save(
            DocumentoRecord(
                id=0,
                carregamento_id=carreg.id,
                usuario_id=usuario_id,
                tipo=DOC_TIPO_MINUTA,
                caminho_arquivo="rebuild/test.pdf",
                nome_arquivo="test.pdf",
                hash_sha256="0" * 64,
            )
        )
        generated["documento"] = doc.id
        assert doc.id > 0

        hist_repo = SqlHistoricoRepository(uow.session)
        hist = hist_repo.append(
            HistoricoRecord(
                id=0,
                carregamento_id=carreg.id,
                usuario_id=usuario_id,
                evento="REBUILD_VALIDATION",
                descricao="insert sem id manual",
            )
        )
        generated["historico_operacional"] = hist.id
        assert hist.id > 0

        uow.session.execute(text("DELETE FROM historico_operacional WHERE id = :id"), {"id": hist.id})
        uow.session.execute(text("DELETE FROM documento WHERE id = :id"), {"id": doc.id})
        uow.session.execute(text("DELETE FROM item_carregamento WHERE carregamento_id = :id"), {"id": carreg.id})
        uow.session.execute(text("DELETE FROM carregamento WHERE id = :id"), {"id": carreg.id})
        uow.session.execute(text("DELETE FROM item_nota_fiscal WHERE nota_fiscal_id = :id"), {"id": nf.id})
        uow.session.execute(text("DELETE FROM nota_fiscal WHERE id = :id"), {"id": nf.id})
        uow.session.execute(text("DELETE FROM configuracao WHERE id = :id"), {"id": cfg.id})

    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    return generated


def rebuild(database_path: Path) -> dict[str, object]:
    if not database_path.exists():
        raise FileNotFoundError(f"Banco legado nao encontrado: {database_path}")

    backups_dir = database_path.parent / "backups"
    backup_path = backup_database(database_path, backups_dir)

    staging_path = database_path.with_suffix(".rebuild.db")
    if staging_path.exists():
        staging_path.unlink()

    create_schema(staging_path)

    source = sqlite3.connect(database_path)
    target = sqlite3.connect(staging_path)
    source.row_factory = sqlite3.Row
    try:
        migrated: dict[str, int] = {}
        for table in MIGRATION_TABLES:
            if table in list_user_tables(source):
                migrated[table] = copy_table_data(source, target, table)
            else:
                migrated[table] = 0

        target.execute("PRAGMA foreign_keys=ON")
        sync_sqlite_sequences(target)
        stamp_alembic(target, ALEMBIC_HEAD)

        pk_errors = validate_integer_primary_keys(target)
        if pk_errors:
            raise RuntimeError("Validacao PK falhou: " + "; ".join(pk_errors))

        fk_errors = validate_foreign_keys(target)
        if fk_errors:
            raise RuntimeError("Validacao FK falhou: " + "; ".join(str(e) for e in fk_errors))

        meta_counts = validate_indexes_and_constraints(target)
    finally:
        source.close()
        target.close()

    autoincrement_ids = validate_autoincrement_inserts(staging_path)

    legacy_path = database_path.with_suffix(".legacy.db")
    if legacy_path.exists():
        legacy_path.unlink()

    try:
        database_path.rename(legacy_path)
        staging_path.rename(database_path)
    except PermissionError as exc:
        raise PermissionError(
            f"Nao foi possivel substituir {database_path}. "
            "Encerre processos que usam o banco (ex.: streamlit run app.py) e execute novamente. "
            f"Banco reconstruido disponivel em: {staging_path}"
        ) from exc

    return {
        "backup_path": str(backup_path),
        "legacy_path": str(legacy_path),
        "new_database_path": str(database_path),
        "migrated_rows": migrated,
        "meta_counts": meta_counts,
        "autoincrement_test_ids": autoincrement_ids,
        "alembic_head": ALEMBIC_HEAD,
    }


def print_report(result: dict[str, object]) -> None:
    print("=== REBUILD SQLITE CONCLUIDO ===")
    for key, value in result.items():
        print(f"{key}: {value}")


def main() -> None:
    settings = get_settings()
    database_path = resolve_sqlite_path(settings.database_url)
    print(f"Reconstruindo: {database_path}")
    result = rebuild(database_path)
    print_report(result)

    conn = sqlite3.connect(database_path)
    try:
        print("\n=== PRAGMA table_info (id) ===")
        for table in TABLES_WITH_INTEGER_PK:
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info([{table}])")
            id_col = next((row for row in cur.fetchall() if row[1] == "id"), None)
            print(f"{table}: type={id_col[2] if id_col else 'N/A'} pk={id_col[5] if id_col else 'N/A'}")
        cur = conn.cursor()
        cur.execute("SELECT version_num FROM alembic_version")
        print(f"\nalembic_version: {cur.fetchone()[0]}")
        cur.execute("SELECT name, seq FROM sqlite_sequence ORDER BY name")
        print("\nsqlite_sequence:")
        for row in cur.fetchall():
            print(row)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
