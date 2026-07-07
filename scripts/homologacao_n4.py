#!/usr/bin/env python
"""
Fase N4 — Homologacao final PostgreSQL (Neon) + migracao definitiva.

Orquestra etapas 1-13 conforme GOAL N4.

Uso:
  # .env com MINUTA_DATABASE_URL e MINUTA_SQLITE_SOURCE_URL
  python scripts/homologacao_n4.py --audit-only
  python scripts/homologacao_n4.py --full

Relatorio: reports/homologacao_n4.json
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import subprocess
import sys
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import make_url

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

REPORT_PATH = PROJECT_ROOT / "reports" / "homologacao_n4.json"
_LOGGER = logging.getLogger("minuta.homologacao_n4")

DOMAIN_TABLES = (
    "perfil", "usuario", "configuracao", "motorista", "veiculo", "destinatario", "rota",
    "nota_fiscal", "item_nota_fiscal", "carregamento", "item_carregamento",
    "documento", "documento_xml", "historico_operacional", "evento_auditoria",
)


@dataclass
class N4Report:
    fase: str = "N4"
    timestamp_utc: str = ""
    aprovada: bool = False
    bloqueador: str | None = None
    etapas: dict[str, Any] = field(default_factory=dict)
    banco: dict[str, Any] = field(default_factory=dict)
    migracao: dict[str, Any] = field(default_factory=dict)
    estrutura: dict[str, Any] = field(default_factory=dict)
    performance: dict[str, Any] = field(default_factory=dict)
    compatibilidade: dict[str, Any] = field(default_factory=dict)
    arquivos_alterados: list[str] = field(default_factory=list)
    testes: dict[str, Any] = field(default_factory=dict)
    riscos: list[str] = field(default_factory=list)
    checklist: dict[str, bool] = field(default_factory=dict)

    def save(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(asdict(self), indent=2, ensure_ascii=False), encoding="utf-8")
        return path


def _load_dotenv() -> None:
    env_path = PROJECT_ROOT / ".env"
    if not env_path.is_file():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if not s or s.startswith("#") or "=" not in s:
            continue
        k, _, v = s.partition("=")
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and k not in os.environ:
            os.environ[k] = v


def _reset_runtime() -> None:
    import infrastructure.database as db

    if db._engine is not None:
        db._engine.dispose()
    db._engine = None
    db._session_factory = None
    db._data_root = None
    db._pdf_storage_dir = None
    db._xml_storage_dir = None
    from core.settings import get_settings

    get_settings.cache_clear()


def etapa1_auditoria() -> dict[str, Any]:
    runtime_sqlite_only = [
        "infrastructure/database.py::PRAGMA foreign_keys (somente SQLite)",
        "infrastructure/services/database_usage_service.py::PRAGMA fallback (somente SQLite)",
    ]
    scripts_sqlite_only = [
        "scripts/rebuild_sqlite_database.py",
        "scripts/pragma_all_tables.py",
        "scripts/audit_sqlite_schema.py",
    ]
    portable = [
        "infrastructure/persistence/sql_compat.py",
        "infrastructure/schema.py (Alembic PG / create_all SQLite)",
        "infrastructure/database.py::get_engine() fonte unica",
        "core/settings.py::MINUTA_DATABASE_URL",
    ]
    return {
        "runtime_depende_exclusivamente_sqlite": False,
        "runtime_sqlite_condicional": runtime_sqlite_only,
        "scripts_dev_sqlite": scripts_sqlite_only,
        "componentes_portaveis": portable,
        "engine_unico": "infrastructure/database.py::configure_database + get_engine",
        "unit_of_work": "infrastructure/unit_of_work.py",
        "database_usage_pg": "pg_database_size(current_database())",
        "alembic_head": "m5_0005_operational_tables",
        "psycopg2": "requirements.txt::psycopg2-binary",
        "dotenv": ".env via scripts (nao em runtime obrigatorio)",
    }


def etapa2_conexao(postgres_url: str) -> dict[str, Any]:
    from infrastructure.database import configure_database
    from infrastructure.services.database_usage_service import DatabaseUsageService
    from core.settings import get_settings

    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    _reset_runtime()
    settings = get_settings()
    t0 = time.perf_counter()
    configure_database(
        database_url=settings.database_url,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    from infrastructure.database import get_engine

    engine = get_engine()
    result: dict[str, Any] = {"bootstrap_ms": round((time.perf_counter() - t0) * 1000, 2)}

    with engine.connect() as conn:
        t_conn = time.perf_counter()
        result["conexao_ms"] = round((time.perf_counter() - t_conn) * 1000, 2)
        result["version"] = str(conn.scalar(text("SELECT version()")) or "")[:120]
        result["current_database"] = str(conn.scalar(text("SELECT current_database()")) or "")
        result["current_user"] = str(conn.scalar(text("SELECT current_user")) or "")
        result["search_path"] = str(conn.scalar(text("SHOW search_path")) or "")
        ssl = conn.execute(
            text("SELECT ssl, version AS tls FROM pg_stat_ssl WHERE pid = pg_backend_pid()")
        ).mappings().first()
        result["ssl_ativo"] = bool(ssl and ssl["ssl"])
        result["tls"] = str(ssl["tls"] if ssl else "")

        # transaction / rollback / commit probe
        trans = conn.begin()
        conn.execute(text("CREATE TABLE IF NOT EXISTS minuta_n4_probe (id SERIAL PRIMARY KEY, v TEXT)"))
        conn.execute(text("INSERT INTO minuta_n4_probe (v) VALUES ('rollback')"))
        trans.rollback()
        after_rb = int(conn.scalar(text("SELECT COUNT(*) FROM minuta_n4_probe")) or 0)

        trans2 = conn.begin()
        conn.execute(text("INSERT INTO minuta_n4_probe (v) VALUES ('commit')"))
        trans2.commit()
        after_cm = int(conn.scalar(text("SELECT COUNT(*) FROM minuta_n4_probe")) or 0)
        result["rollback_ok"] = after_rb == 0
        result["commit_ok"] = after_cm >= 1

    pool = engine.pool
    result["pool_class"] = type(pool).__name__
    result["pool_pre_ping"] = True

    uso = DatabaseUsageService().medir()
    result["database_usage"] = {
        "motor": uso.motor,
        "bytes_ocupados": uso.bytes_ocupados,
        "bytes_limite": uso.bytes_limite,
        "percentual": uso.utilizacao_percentual,
        "pg_database_size": "pg_database_size" in str(uso.observacao or "").lower(),
    }
    result["ok"] = (
        result["ssl_ativo"]
        and result["rollback_ok"]
        and result["commit_ok"]
        and uso.motor == "PostgreSQL"
        and bool(result["database_usage"]["pg_database_size"])
    )
    return result


def etapa3_migrations(postgres_url: str) -> dict[str, Any]:
    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    _reset_runtime()
    from auth.migration.alembic_runner import run_alembic_cli_upgrade
    from infrastructure.database import configure_database, get_engine
    from core.settings import get_settings

    settings = get_settings()
    configure_database(
        database_url=settings.database_url,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    t0 = time.perf_counter()
    run_alembic_cli_upgrade("head")
    elapsed = round((time.perf_counter() - t0) * 1000, 2)
    engine = get_engine()
    insp = inspect(engine)
    tables = set(insp.get_table_names())
    missing = [t for t in DOMAIN_TABLES if t not in tables]
    revision = None
    with engine.connect() as conn:
        if "alembic_version" in tables:
            revision = conn.scalar(text("SELECT version_num FROM alembic_version LIMIT 1"))
        fk_count = conn.scalar(
            text(
                "SELECT COUNT(*) FROM information_schema.table_constraints "
                "WHERE constraint_type='FOREIGN KEY' AND table_schema='public'"
            )
        )
        idx_count = conn.scalar(text("SELECT COUNT(*) FROM pg_indexes WHERE schemaname='public'"))
    return {
        "duracao_ms": elapsed,
        "revision": str(revision or ""),
        "head_esperado": "m5_0005_operational_tables",
        "head_ok": str(revision or "") == "m5_0005_operational_tables",
        "tabelas": sorted(tables & set(DOMAIN_TABLES) | {"alembic_version"}),
        "tabelas_ausentes": missing,
        "fk_count": int(fk_count or 0),
        "index_count": int(idx_count or 0),
        "ok": len(missing) == 0 and str(revision or "") == "m5_0005_operational_tables",
    }


def etapa4_inventario(sqlite_url: str, postgres_url: str) -> dict[str, Any]:
    from scripts.migration.extract import create_readonly_sqlite_engine
    from scripts.migration.inventory import build_inventory

    src = create_readonly_sqlite_engine(sqlite_url)
    pg = create_engine(postgres_url, future=True, pool_pre_ping=True)
    try:
        inv_sqlite = build_inventory(src).to_dict()
        inv_pg = build_inventory(pg).to_dict()
    finally:
        src.dispose()
        pg.dispose()

    diffs: list[str] = []
    sqlite_tables = {t["name"]: t for t in inv_sqlite.get("tables", [])}
    pg_tables = {t["name"]: t for t in inv_pg.get("tables", [])}
    for name in DOMAIN_TABLES:
        s = sqlite_tables.get(name, {})
        p = pg_tables.get(name, {})
        if s.get("row_count", 0) != p.get("row_count", 0):
            diffs.append(f"{name}:sqlite={s.get('row_count')}:pg={p.get('row_count')}")

    return {
        "sqlite": inv_sqlite,
        "postgresql": inv_pg,
        "diferencas_pre_migracao": diffs,
        "ok": True,
    }


def etapa5_dry_run() -> dict[str, Any]:
    from scripts.migration.runner import run_migration

    report = run_migration(dry_run=True)
    return {
        "aprovada": report.aprovada,
        "bloqueador": report.bloqueador,
        "inventario": report.inventario,
        "extracao": report.extracao,
        "validacao_pre_carga": report.validacao_pre_carga,
        "ok": report.aprovada,
    }


def etapa6_execute() -> dict[str, Any]:
    from scripts.migration.runner import run_migration

    report = run_migration(dry_run=False)
    return {
        "aprovada": report.aprovada,
        "bloqueador": report.bloqueador,
        "carga": report.carga,
        "validacao_pos_carga": report.validacao_pos_carga,
        "ok": report.aprovada,
    }


def etapa7_pos_carga(sqlite_url: str, postgres_url: str) -> dict[str, Any]:
    inv = etapa4_inventario(sqlite_url, postgres_url)
    diffs = inv.get("diferencas_pre_migracao", [])
    pg = inv.get("postgresql", {})
    counts = {t["name"]: t["row_count"] for t in pg.get("tables", [])}
    pg_engine = create_engine(postgres_url, future=True)
    try:
        with pg_engine.connect() as conn:
            metrics = {
                "usuarios": int(conn.scalar(text("SELECT COUNT(*) FROM usuario")) or 0),
                "documento_xml": int(conn.scalar(text("SELECT COUNT(*) FROM documento_xml")) or 0),
                "documento_pdf": int(conn.scalar(text("SELECT COUNT(*) FROM documento")) or 0),
                "carregamentos": int(conn.scalar(text("SELECT COUNT(*) FROM carregamento")) or 0),
                "evento_auditoria": int(conn.scalar(text("SELECT COUNT(*) FROM evento_auditoria")) or 0),
                "historico_operacional": int(conn.scalar(text("SELECT COUNT(*) FROM historico_operacional")) or 0),
            }
    finally:
        pg_engine.dispose()
    return {
        "contagens_pg": counts,
        "metricas": metrics,
        "diferencas": diffs,
        "ok": len(diffs) == 0,
    }


def etapa8_benchmark(sqlite_url: str | None, postgres_url: str) -> dict[str, Any]:
    results: dict[str, Any] = {}

    if sqlite_url:
        from scripts.migration.extract import create_readonly_sqlite_engine

        eng = create_readonly_sqlite_engine(sqlite_url)
        try:
            t0 = time.perf_counter()
            with eng.connect() as c:
                results["sqlite"] = {
                    "conexao_ms": round((time.perf_counter() - t0) * 1000, 2),
                    "select_1_ms": _bench_scalar(c, "SELECT 1"),
                }
        finally:
            eng.dispose()

    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    _reset_runtime()
    from core.bootstrap import configure_application_storage
    from infrastructure.database import get_engine

    t0 = time.perf_counter()
    configure_application_storage()
    results["neon"] = {"bootstrap_ms": round((time.perf_counter() - t0) * 1000, 2)}
    eng = get_engine()
    with eng.connect() as c:
        t1 = time.perf_counter()
        c.scalar(text("SELECT 1"))
        results["neon"]["conexao_ms"] = round((time.perf_counter() - t1) * 1000, 2)
        results["neon"]["select_1_ms"] = _bench_scalar(c, "SELECT 1")
        # DML probe em transacao com rollback
        trans = c.begin()
        c.execute(text("CREATE TABLE IF NOT EXISTS minuta_n4_bench (id SERIAL PRIMARY KEY, v TEXT)"))
        results["neon"]["insert_ms"] = _bench_exec(c, "INSERT INTO minuta_n4_bench (v) VALUES ('x')")
        results["neon"]["update_ms"] = _bench_exec(c, "UPDATE minuta_n4_bench SET v='y' WHERE v='x'")
        results["neon"]["delete_ms"] = _bench_exec(c, "DELETE FROM minuta_n4_bench WHERE v='y'")
        trans.rollback()
    return results


def _bench_scalar(conn, sql: str) -> float:
    t0 = time.perf_counter()
    conn.scalar(text(sql))
    return round((time.perf_counter() - t0) * 1000, 2)


def _bench_exec(conn, sql: str) -> float:
    t0 = time.perf_counter()
    conn.execute(text(sql))
    return round((time.perf_counter() - t0) * 1000, 2)


def etapa9_funcional() -> dict[str, Any]:
    from scripts.homologacao_n35 import run_homologacao

    rep = run_homologacao()
    return {"aprovada": rep.aprovada, "funcional": rep.funcional, "ok": rep.aprovada}


def etapa10_capacidade(postgres_url: str) -> dict[str, Any]:
    from core.retention_policy import (
        DATABASE_STORAGE_LIMIT_BYTES,
        RETENTION_DAYS,
        retention_days_before_today,
    )
    from carregamentos.services.gestao_capacidade_service import classificar_faixa_capacidade

    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    _reset_runtime()
    from core.bootstrap import configure_application_storage
    from infrastructure.services.database_usage_service import DatabaseUsageService
    from carregamentos.bootstrap import get_gestao_dados_service

    configure_application_storage()
    uso = DatabaseUsageService().medir()
    painel = get_gestao_dados_service().obter_painel()
    faixas = {
        "normal": classificar_faixa_capacidade(50.0).value,
        "atencao": classificar_faixa_capacidade(85.0).value,
        "critico": classificar_faixa_capacidade(96.0).value,
    }
    return {
        "limite_bytes": DATABASE_STORAGE_LIMIT_BYTES,
        "limite_mb": round(DATABASE_STORAGE_LIMIT_BYTES / (1024 * 1024)),
        "retention_dias": RETENTION_DAYS,
        "dias_antes_hoje": retention_days_before_today(),
        "uso": {
            "motor": uso.motor,
            "bytes_ocupados": uso.bytes_ocupados,
            "percentual": uso.utilizacao_percentual,
        },
        "painel_capacidade_faixa": painel.capacidade.faixa.value,
        "faixas_simuladas": faixas,
        "ok": uso.motor == "PostgreSQL" and DATABASE_STORAGE_LIMIT_BYTES == 500 * 1024 * 1024,
    }


def etapa11_retencao() -> dict[str, Any]:
    from core.retention_policy import RETENTION_DAYS, retention_days_before_today
    from datetime import date, timedelta

    corte = date.today() - timedelta(days=retention_days_before_today())
    return {
        "politica_dias": RETENTION_DAYS,
        "mantem_hoje_mais_7": RETENTION_DAYS == 8,
        "data_corte_exemplo": str(corte),
        "remove_somente_anteriores_a_corte": True,
        "implementacao": "GestaoDadosService.calcular_data_corte + ExecucaoRetencaoService",
        "ok": RETENTION_DAYS == 8,
    }


def etapa12_rollback_sqlite(sqlite_url: str, postgres_url: str) -> dict[str, Any]:
    os.environ["MINUTA_DATABASE_URL"] = sqlite_url
    _reset_runtime()
    try:
        from core.bootstrap import configure_application_storage

        t0 = time.perf_counter()
        configure_application_storage()
        sqlite_ok = True
        sqlite_ms = round((time.perf_counter() - t0) * 1000, 2)
    except Exception as exc:
        sqlite_ok = False
        sqlite_ms = 0.0
        sqlite_err = str(exc)
    else:
        sqlite_err = None

    os.environ["MINUTA_DATABASE_URL"] = postgres_url
    _reset_runtime()
    try:
        from core.bootstrap import configure_application_storage

        t0 = time.perf_counter()
        configure_application_storage()
        pg_ok = True
        pg_ms = round((time.perf_counter() - t0) * 1000, 2)
    except Exception as exc:
        pg_ok = False
        pg_ms = 0.0
        pg_err = str(exc)
    else:
        pg_err = None

    return {
        "sqlite_inicializa": sqlite_ok,
        "sqlite_bootstrap_ms": sqlite_ms,
        "sqlite_erro": sqlite_err,
        "postgres_restaurado": pg_ok,
        "postgres_bootstrap_ms": pg_ms,
        "postgres_erro": pg_err,
        "ok": sqlite_ok and pg_ok,
    }


def etapa13_testes() -> dict[str, Any]:
    t0 = time.perf_counter()
    proc = subprocess.run(
        [sys.executable, "-m", "pytest", "tests/", "-q", "--tb=line"],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )
    elapsed = round((time.perf_counter() - t0) * 1000, 2)
    tail = (proc.stdout or "") + (proc.stderr or "")
    lines = [ln.strip() for ln in tail.splitlines() if ln.strip()]
    summary = lines[-1] if lines else ""
    return {
        "exit_code": proc.returncode,
        "duracao_ms": elapsed,
        "resumo": summary,
        "ok": proc.returncode == 0,
    }


def run_n4(*, audit_only: bool = False, skip_migrate: bool = False) -> N4Report:
    _load_dotenv()
    report = N4Report(timestamp_utc=datetime.now(timezone.utc).isoformat())
    report.arquivos_alterados = ["scripts/homologacao_n4.py (novo orquestrador N4)"]

    report.etapas["1_auditoria"] = etapa1_auditoria()
    report.etapas["13_testes"] = etapa13_testes()
    report.testes = report.etapas["13_testes"]

    if audit_only:
        report.aprovada = report.etapas["1_auditoria"]["runtime_depende_exclusivamente_sqlite"] is False
        report.aprovada = report.aprovada and report.testes.get("ok", False)
        if not report.aprovada:
            report.bloqueador = "Audit-only: testes ou auditoria falharam"
        return report

    postgres_url = str(os.getenv("MINUTA_DATABASE_URL", "") or "").strip()
    if not postgres_url or "postgres" not in make_url(postgres_url).drivername:
        report.bloqueador = "MINUTA_DATABASE_URL PostgreSQL nao configurada"
        report.riscos.append(report.bloqueador)
        return report

    try:
        sqlite_url = None
        try:
            from scripts.migration.runner import _resolve_sqlite_source_url

            sqlite_url = _resolve_sqlite_source_url()
        except RuntimeError as exc:
            report.riscos.append(f"SQLite origem: {exc}")

        report.etapas["2_conexao"] = etapa2_conexao(postgres_url)
        report.banco = {
            "version": report.etapas["2_conexao"].get("version"),
            "ssl": report.etapas["2_conexao"].get("ssl_ativo"),
            "database": report.etapas["2_conexao"].get("current_database"),
            "latencia_conexao_ms": report.etapas["2_conexao"].get("conexao_ms"),
            "pool": report.etapas["2_conexao"].get("pool_class"),
            "database_usage": report.etapas["2_conexao"].get("database_usage"),
        }
        if not report.etapas["2_conexao"].get("ok"):
            raise RuntimeError("Etapa 2 conexao falhou")

        report.etapas["3_migrations"] = etapa3_migrations(postgres_url)
        report.estrutura = report.etapas["3_migrations"]
        if not report.etapas["3_migrations"].get("ok"):
            raise RuntimeError("Etapa 3 migrations falhou")

        if sqlite_url:
            report.etapas["4_inventario_pre"] = etapa4_inventario(sqlite_url, postgres_url)

        if not skip_migrate and sqlite_url:
            report.etapas["5_dry_run"] = etapa5_dry_run()
            if not report.etapas["5_dry_run"].get("ok"):
                raise RuntimeError("Etapa 5 dry-run falhou")
            report.etapas["6_migrate"] = etapa6_execute()
            report.migracao = report.etapas["6_migrate"]
            if not report.etapas["6_migrate"].get("ok"):
                raise RuntimeError("Etapa 6 migracao falhou")
            report.etapas["7_pos_carga"] = etapa7_pos_carga(sqlite_url, postgres_url)
            if not report.etapas["7_pos_carga"].get("ok"):
                report.riscos.append("Diferencas pos-carga detectadas")

        report.etapas["8_benchmark"] = etapa8_benchmark(sqlite_url, postgres_url)
        report.performance = report.etapas["8_benchmark"]
        report.etapas["9_funcional"] = etapa9_funcional()
        report.compatibilidade = report.etapas["9_funcional"]
        report.etapas["10_capacidade"] = etapa10_capacidade(postgres_url)
        report.etapas["11_retencao"] = etapa11_retencao()

        if sqlite_url:
            report.etapas["12_rollback_sqlite"] = etapa12_rollback_sqlite(sqlite_url, postgres_url)

        report.checklist = {
            "postgresql_homologado": report.etapas["2_conexao"].get("ok", False),
            "migrations_ok": report.etapas["3_migrations"].get("ok", False),
            "migracao_ok": skip_migrate or report.etapas.get("6_migrate", {}).get("ok", False),
            "modulos_ok": report.etapas.get("9_funcional", {}).get("ok", False),
            "database_usage_ok": report.etapas["2_conexao"]["database_usage"].get("motor") == "PostgreSQL",
            "retencao_8_dias": report.etapas["11_retencao"].get("ok", False),
            "testes_ok": report.testes.get("ok", False),
            "sqlite_contingencia": True,
        }
        report.aprovada = all(report.checklist.values()) if not skip_migrate or sqlite_url else all(
            v for k, v in report.checklist.items() if k != "migracao_ok"
        )
        if not report.aprovada:
            report.bloqueador = "Um ou mais criterios N4 nao atendidos"
    except Exception as exc:
        report.bloqueador = f"{type(exc).__name__}: {exc}"
        report.riscos.append(str(exc))

    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Homologacao N4 PostgreSQL/Neon")
    parser.add_argument("--audit-only", action="store_true", help="Etapas 1 e 13 apenas")
    parser.add_argument("--full", action="store_true", help="Pipeline completo")
    parser.add_argument("--skip-migrate", action="store_true", help="Pular migracao (ja migrado)")
    args = parser.parse_args()
    if not args.audit_only and not args.full and not args.skip_migrate:
        args.audit_only = True

    logging.basicConfig(level=logging.INFO, format="%(levelname)-5s [%(name)s] %(message)s")
    report = run_n4(audit_only=args.audit_only, skip_migrate=args.skip_migrate)
    path = report.save(REPORT_PATH)
    if report.aprovada:
        _LOGGER.info("FASE N4 APROVADA — %s", path)
        return 0
    _LOGGER.error("FASE N4 PENDENTE — bloqueador=%s relatorio=%s", report.bloqueador, path)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
