from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine, make_url

from core.settings import ensure_dotenv_loaded, get_settings, reset_settings_cache
from scripts.migration.constants import DOMAIN_TABLES
from scripts.migration.extract import create_readonly_sqlite_engine, extract_all
from scripts.migration.inventory import build_inventory, validate_source_integrity
from scripts.migration.load import load_dataset
from scripts.migration.report import MigrationReport, compare_inventories

_LOGGER = logging.getLogger("minuta.migration.runner")
PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _resolve_sqlite_source_url() -> str:
    settings = get_settings()
    if settings.sqlite_source_url:
        return settings.sqlite_source_url
    default_path = (PROJECT_ROOT / "data" / "minuta.db").resolve()
    if default_path.is_file():
        return f"sqlite:///{default_path.as_posix()}"
    raise RuntimeError(
        "Fonte SQLite nao encontrada. Defina MINUTA_SQLITE_SOURCE_URL ou crie data/minuta.db"
    )


def _resolve_postgres_target_url() -> str:
    settings = get_settings()
    target = str(settings.database_url or "").strip()
    if not target:
        raise RuntimeError("MINUTA_DATABASE_URL nao definida (destino PostgreSQL/Neon).")
    driver = make_url(target).drivername or ""
    if "postgres" not in driver:
        raise RuntimeError("MINUTA_DATABASE_URL deve apontar para PostgreSQL/Neon.")
    return target


def _mask_config(url: str) -> dict[str, Any]:
    parsed = make_url(url)
    return {
        "driver": parsed.drivername,
        "host": parsed.host,
        "database": parsed.database,
        "dialect": "postgresql" if "postgres" in (parsed.drivername or "") else parsed.drivername,
    }


def _ensure_target_schema(postgres_url: str) -> None:
    from auth.migration.alembic_runner import run_alembic_cli_upgrade

    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    reset_settings_cache()
    run_alembic_cli_upgrade("head")


def _count_files(data_root: Path) -> dict[str, int]:
    xml_dir = data_root / "xml_storage"
    pdf_dirs = [data_root / "documentos", data_root / "carregamentos"]
    xml_count = sum(1 for _ in xml_dir.rglob("*.xml")) if xml_dir.is_dir() else 0
    pdf_count = 0
    for directory in pdf_dirs:
        if directory.is_dir():
            pdf_count += sum(1 for _ in directory.rglob("*.pdf"))
    return {"xml_arquivos": xml_count, "pdf_arquivos": pdf_count}


def _target_has_data(engine: Engine) -> bool:
    from sqlalchemy import inspect as sa_inspect

    inspector = sa_inspect(engine)
    with engine.connect() as connection:
        for table_name in DOMAIN_TABLES:
            if table_name not in inspector.get_table_names():
                continue
            count = int(connection.scalar(text(f'SELECT COUNT(*) FROM "{table_name}"')) or 0)
            if count > 0:
                return True
    return False


def run_migration(*, dry_run: bool = False) -> MigrationReport:
    ensure_dotenv_loaded()
    reset_settings_cache()
    settings = get_settings()
    report = MigrationReport()

    sqlite_url = _resolve_sqlite_source_url()
    report.sqlite_source = _mask_config(sqlite_url)

    batch_size = settings.migration_batch_size
    data_root = settings.data_root
    report.arquivos = _count_files(data_root)

    source_engine = create_readonly_sqlite_engine(sqlite_url)

    try:
        inv_start = time.perf_counter()
        inventory = build_inventory(source_engine)
        report.inventario = inventory.to_dict()
        report.tempos_ms["inventario"] = round((time.perf_counter() - inv_start) * 1000, 2)

        pre_errors = validate_source_integrity(source_engine, inventory)
        report.validacao_pre_carga = {"ok": len(pre_errors) == 0, "erros": pre_errors}
        if pre_errors:
            report.bloqueador = "Validacao pre-carga falhou"
            report.erros.extend(pre_errors)
            raise RuntimeError("Abortado: inconsistencias na origem SQLite")

        ext_start = time.perf_counter()
        extraction = extract_all(source_engine)
        report.extracao = extraction.to_dict()
        report.tempos_ms["extracao"] = extraction.duration_ms

        if dry_run:
            report.avisos.append("Dry-run: nenhum dado carregado no PostgreSQL.")
            report.aprovada = True
            return report

        postgres_url = _resolve_postgres_target_url()
        report.postgres_target = _mask_config(postgres_url)
        target_engine = create_engine(postgres_url, future=True, pool_pre_ping=True)

        try:
            schema_start = time.perf_counter()
            _ensure_target_schema(postgres_url)
            report.tempos_ms["alembic"] = round((time.perf_counter() - schema_start) * 1000, 2)

            if _target_has_data(target_engine):
                report.avisos.append(
                    "Destino continha dados — TRUNCATE CASCADE sera executado antes da carga."
                )

            load_start = time.perf_counter()
            load_report = load_dataset(
                target_engine, extraction.tables, batch_size=batch_size, truncate=True
            )
            report.carga = load_report.to_dict()
            report.tempos_ms["carga"] = load_report.duration_ms

            post_start = time.perf_counter()
            target_inventory = build_inventory(target_engine)
            comparison = compare_inventories(report.inventario, target_inventory.to_dict())
            report.validacao_pos_carga = comparison
            report.tempos_ms["validacao_pos_carga"] = round((time.perf_counter() - post_start) * 1000, 2)
            if comparison.get("avisos_checksum"):
                report.avisos.extend(comparison["avisos_checksum"])

            if not comparison["equivalente"]:
                report.bloqueador = "Validacao pos-carga divergente"
                report.erros.extend(comparison["diferencas"])
                raise RuntimeError("Abortado: contagens divergentes apos carga")

            report.aprovada = True
            report.rollback = {
                "estrategia": "transacao_com_truncate_e_rollback_em_falha",
                "sqlite_preservado": True,
                "instrucao": "Para rollback operacional, redefina MINUTA_DATABASE_URL para o SQLite de origem.",
            }
            return report
        finally:
            target_engine.dispose()

    except Exception as exc:
        if not report.bloqueador:
            report.bloqueador = f"{type(exc).__name__}: {exc}"
        report.rollback = {
            "executado": True,
            "sqlite_preservado": True,
            "observacao": "PostgreSQL nao deve conter carga parcial (rollback transacional).",
        }
        return report
    finally:
        source_engine.dispose()
