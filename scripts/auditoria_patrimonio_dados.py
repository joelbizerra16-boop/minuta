#!/usr/bin/env python
"""
Auditoria completa do patrimonio de dados (SOMENTE LEITURA).

Gera:
  reports/auditoria_patrimonio_dados.json
  reports/auditoria_patrimonio_dados.md

Nao altera banco, arquivos ou regras de negocio.
"""

from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
REPORT_JSON = PROJECT_ROOT / "reports" / "auditoria_patrimonio_dados.json"
REPORT_MD = PROJECT_ROOT / "reports" / "auditoria_patrimonio_dados.md"

DOMAIN_TABLES = (
    "perfil", "usuario", "configuracao", "motorista", "veiculo", "destinatario", "rota",
    "nota_fiscal", "item_nota_fiscal", "carregamento", "item_carregamento",
    "documento", "documento_xml", "historico_operacional", "evento_auditoria",
)

XML_DIR_NAMES = {"xml_storage", "xml", "uploads", "storage"}
PDF_DIR_NAMES = {"documentos", "carregamentos", "pdf", "pdfs", "uploads", "storage"}


@dataclass
class RiskItem:
    nivel: str
    codigo: str
    descricao: str
    impacto: str


@dataclass
class AuditReport:
    timestamp_utc: str = ""
    sqlite_path: str | None = None
    sqlite_found: bool = False
    apto_migracao: bool = False
    bloqueador: str | None = None
    fase_a_tabelas: list[dict[str, Any]] = field(default_factory=list)
    fase_b_documento_xml: dict[str, Any] = field(default_factory=dict)
    fase_c_documento: dict[str, Any] = field(default_factory=dict)
    fase_d_notas: dict[str, Any] = field(default_factory=dict)
    fase_e_diretorios: list[dict[str, Any]] = field(default_factory=list)
    fase_f_xml_consistencia: dict[str, Any] = field(default_factory=dict)
    fase_g_pdf_consistencia: dict[str, Any] = field(default_factory=dict)
    fase_h_espaco: dict[str, Any] = field(default_factory=dict)
    fase_i_integridade: dict[str, Any] = field(default_factory=dict)
    fase_j_neon_prep: dict[str, Any] = field(default_factory=dict)
    fase_k_riscos: list[dict[str, Any]] = field(default_factory=list)
    resumo: dict[str, Any] = field(default_factory=dict)
    tempos_ms: dict[str, float] = field(default_factory=dict)


def _load_dotenv_readonly() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    env_path = PROJECT_ROOT / ".env"
    if env_path.is_file():
        load_dotenv(env_path, override=False)


def _sqlite_path_from_url(url: str) -> Path | None:
    if not url or not url.startswith("sqlite:///"):
        return None
    raw = url.replace("sqlite:///", "", 1)
    if raw == ":memory:":
        return None
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    return path


def _discover_sqlite_path() -> tuple[Path | None, list[str]]:
    attempts: list[str] = []
    _load_dotenv_readonly()

    candidates: list[Path] = []
    for env_name in ("MINUTA_SQLITE_SOURCE_URL", "MINUTA_DATABASE_URL"):
        url = str(os.getenv(env_name, "") or "").strip()
        if url.startswith("sqlite:///"):
            p = _sqlite_path_from_url(url)
            if p:
                candidates.append(p)
                attempts.append(f"{env_name}={p}")

    default_path = (PROJECT_ROOT / "data" / "minuta.db").resolve()
    candidates.append(default_path)
    attempts.append(f"default={default_path}")

    search_roots = [PROJECT_ROOT, PROJECT_ROOT.parent]
    for root in search_roots:
        if not root.exists():
            continue
        for pattern in ("minuta.db", "minuta_*.db"):
            for match in root.rglob(pattern):
                s = str(match)
                if any(x in s for x in (".venv", "pytest", "node_modules", "__pycache__")):
                    continue
                candidates.append(match.resolve())

    seen: set[str] = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        if path.is_file():
            return path, attempts

    return None, attempts


def _open_sqlite_readonly(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{db_path.as_posix()}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA query_only=ON")
    return conn


def _table_names(conn: sqlite3.Connection) -> list[str]:
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    ).fetchall()
    return [str(r[0]) for r in rows]


def _table_info(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    rows = conn.execute(f"PRAGMA table_info([{table}])").fetchall()
    return [
        {
            "cid": r[0],
            "name": r[1],
            "type": r[2],
            "notnull": bool(r[3]),
            "default": r[4],
            "pk": bool(r[5]),
        }
        for r in rows
    ]


def _table_indexes(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    rows = conn.execute(f"PRAGMA index_list([{table}])").fetchall()
    result = []
    for row in rows:
        idx_name = row[1]
        cols = [r[2] for r in conn.execute(f"PRAGMA index_info([{idx_name}])").fetchall()]
        result.append({"name": idx_name, "unique": bool(row[2]), "columns": cols})
    return result


def _table_fks(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    rows = conn.execute(f"PRAGMA foreign_key_list([{table}])").fetchall()
    return [
        {
            "id": r[0],
            "table": r[2],
            "from": r[3],
            "to": r[4],
            "on_update": r[5],
            "on_delete": r[6],
        }
        for r in rows
    ]


def _estimate_table_mb(conn: sqlite3.Connection, table: str) -> float:
    try:
        page_count = conn.execute(
            "SELECT COUNT(*) FROM dbstat WHERE name=? AND aggregate=TRUE", (table,)
        ).fetchone()
        if page_count and page_count[0]:
            page_size = int(conn.execute("PRAGMA page_size").fetchone()[0])
            return round((page_count[0] * page_size) / (1024 * 1024), 4)
    except sqlite3.Error:
        pass
    row = conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()
    count = int(row[0] or 0) if row else 0
    return round((count * 256) / (1024 * 1024), 4)


def _id_range(conn: sqlite3.Connection, table: str, info: list[dict[str, Any]]) -> tuple[Any, Any]:
    pk_cols = [c["name"] for c in info if c["pk"]]
    if not pk_cols:
        return None, None
    col = pk_cols[0]
    try:
        first = conn.execute(f'SELECT MIN("{col}") FROM "{table}"').fetchone()[0]
        last = conn.execute(f'SELECT MAX("{col}") FROM "{table}"').fetchone()[0]
        return first, last
    except sqlite3.Error:
        return None, None


def audit_tables(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    tables = _table_names(conn)
    result = []
    for table in tables:
        info = _table_info(conn, table)
        count = int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
        first_id, last_id = _id_range(conn, table, info)
        result.append(
            {
                "tabela": table,
                "qtd_registros": count,
                "primeiro_id": first_id,
                "ultimo_id": last_id,
                "mb_aproximado": _estimate_table_mb(conn, table),
                "colunas": len(info),
                "primary_key": [c["name"] for c in info if c["pk"]],
                "foreign_keys": _table_fks(conn, table),
                "indices": _table_indexes(conn, table),
            }
        )
    return result


def _audit_table_detail(conn: sqlite3.Connection, table: str, sample: int = 5) -> dict[str, Any]:
    info = _table_info(conn, table)
    count = int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
    rows = conn.execute(f'SELECT * FROM "{table}" ORDER BY rowid LIMIT {sample}').fetchall()
    samples = [dict(r) for r in rows]
    columns = {c["name"]: c for c in info}
    return {
        "quantidade": count,
        "estrutura": info,
        "colunas_tipos": {c["name"]: c["type"] for c in info},
        "primary_key": [c["name"] for c in info if c["pk"]],
        "foreign_keys": _table_fks(conn, table),
        "indices": _table_indexes(conn, table),
        "colunas_esperadas": {
            "xml_blob": any("blob" in c["type"].lower() for c in info),
            "caminho_arquivo": "caminho_arquivo" in columns,
            "hash": "hash_sha256" in columns,
            "tamanho": "tamanho" in columns,
            "chave_nfe": "chave_nfe" in columns,
            "data_importacao": "data_importacao" in columns,
            "data_criacao": "criado_em" in columns or "data_criacao" in columns,
            "nome_arquivo": "nome_arquivo" in columns,
            "tipo": "tipo" in columns,
            "carregamento_id": "carregamento_id" in columns,
        },
        "exemplos": samples,
    }


def _discover_asset_dirs() -> list[Path]:
    found: set[Path] = set()
    search_roots = [
        PROJECT_ROOT,
        PROJECT_ROOT / "data",
        PROJECT_ROOT.parent,
        Path(os.getenv("MINUTA_DATA_ROOT", "") or "").expanduser() if os.getenv("MINUTA_DATA_ROOT") else None,
    ]
    for root in search_roots:
        if root is None or not root.exists():
            continue
        for path in root.rglob("*"):
            if not path.is_dir():
                continue
            name = path.name.lower()
            if name in XML_DIR_NAMES or name in PDF_DIR_NAMES:
                found.add(path.resolve())
            if path.name.lower() == "data" and (path / "xml_storage").is_dir():
                found.add((path / "xml_storage").resolve())
            if path.name.lower() == "data" and (path / "documentos").is_dir():
                found.add((path / "documentos").resolve())

    for path in PROJECT_ROOT.rglob("*.xml"):
        if ".venv" not in str(path) and "pytest" not in str(path):
            found.add(path.parent.resolve())
    for path in PROJECT_ROOT.rglob("*.pdf"):
        if ".venv" not in str(path) and "pytest" not in str(path):
            found.add(path.parent.resolve())

    return sorted(found)


def _scan_directory(directory: Path, extensions: set[str]) -> dict[str, Any]:
    files = [p for p in directory.rglob("*") if p.is_file() and p.suffix.lower() in extensions]
    if not files:
        return {
            "caminho": str(directory),
            "quantidade": 0,
            "bytes": 0,
            "mb": 0.0,
            "maior_arquivo": None,
            "menor_arquivo": None,
        }
    sizes = [(p, p.stat().st_size) for p in files]
    total = sum(s for _, s in sizes)
    largest = max(sizes, key=lambda x: x[1])
    smallest = min(sizes, key=lambda x: x[1])
    return {
        "caminho": str(directory),
        "quantidade": len(files),
        "bytes": total,
        "mb": round(total / (1024 * 1024), 4),
        "maior_arquivo": {"caminho": str(largest[0]), "bytes": largest[1]},
        "menor_arquivo": {"caminho": str(smallest[0]), "bytes": smallest[1]},
    }


def _file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _resolve_file_path(stored: str, data_root: Path) -> Path | None:
    raw = str(stored or "").strip()
    if not raw:
        return None
    path = Path(raw).expanduser()
    if path.is_file():
        return path.resolve()
    candidate = (data_root / raw).resolve()
    if candidate.is_file():
        return candidate
    candidate = (PROJECT_ROOT / raw).resolve()
    if candidate.is_file():
        return candidate
    return None


def _compare_xml(conn: sqlite3.Connection, data_root: Path) -> dict[str, Any]:
    rows = conn.execute(
        'SELECT id, chave_nfe, nome_arquivo, caminho_arquivo, hash_sha256, ativo FROM documento_xml'
    ).fetchall()
    db_by_chave = {str(r["chave_nfe"]): dict(r) for r in rows}
    fs_xml: dict[str, Path] = {}
    for directory in _discover_asset_dirs():
        for path in directory.rglob("*.xml"):
            key = path.stem
            if len(key) == 44 and key.isdigit():
                fs_xml[key] = path.resolve()

    registrado_existente = []
    registrado_inexistente = []
    hash_divergente = []
    nome_divergente = []
    for chave, row in db_by_chave.items():
        path = _resolve_file_path(row["caminho_arquivo"], data_root)
        if path is None:
            alt = fs_xml.get(chave)
            if alt and alt.is_file():
                registrado_existente.append({"chave_nfe": chave, "caminho_db": row["caminho_arquivo"], "caminho_fs": str(alt)})
            else:
                registrado_inexistente.append({"chave_nfe": chave, "caminho_db": row["caminho_arquivo"]})
            continue
        registrado_existente.append({"chave_nfe": chave, "caminho": str(path)})
        if path.name != row["nome_arquivo"]:
            nome_divergente.append({"chave_nfe": chave, "db": row["nome_arquivo"], "fs": path.name})
        try:
            if row["hash_sha256"] and _file_sha256(path) != row["hash_sha256"]:
                hash_divergente.append({"chave_nfe": chave, "caminho": str(path)})
        except OSError:
            registrado_inexistente.append({"chave_nfe": chave, "caminho_db": row["caminho_arquivo"], "erro": "leitura"})

    fs_sem_registro = [{"chave_nfe": k, "caminho": str(v)} for k, v in fs_xml.items() if k not in db_by_chave]
    duplicados_fs = []
    seen: dict[str, list[str]] = {}
    for k, v in fs_xml.items():
        seen.setdefault(k, []).append(str(v))
    duplicados_fs = [{"chave_nfe": k, "caminhos": paths} for k, paths in seen.items() if len(paths) > 1]

    return {
        "registros_db": len(rows),
        "arquivos_fs_chave": len(fs_xml),
        "registrado_e_existente": len(registrado_existente),
        "registrado_inexistente": registrado_inexistente[:100],
        "registrado_inexistente_total": len(registrado_inexistente),
        "existente_sem_registro": fs_sem_registro[:100],
        "existente_sem_registro_total": len(fs_sem_registro),
        "hash_divergente": hash_divergente[:100],
        "hash_divergente_total": len(hash_divergente),
        "nome_divergente": nome_divergente[:100],
        "nome_divergente_total": len(nome_divergente),
        "duplicados_fs": duplicados_fs[:50],
        "duplicados_fs_total": len(duplicados_fs),
    }


def _compare_pdf(conn: sqlite3.Connection, data_root: Path) -> dict[str, Any]:
    rows = conn.execute(
        'SELECT id, carregamento_id, tipo, nome_arquivo, caminho_arquivo, hash_sha256 FROM documento'
    ).fetchall()
    fs_pdf: list[Path] = []
    for directory in _discover_asset_dirs():
        fs_pdf.extend(p.resolve() for p in directory.rglob("*.pdf") if p.is_file())
    fs_by_path = {str(p).lower(): p for p in fs_pdf}

    registrado_existente = 0
    registrado_inexistente = []
    hash_divergente = []
    nome_divergente = []
    for row in rows:
        path = _resolve_file_path(row["caminho_arquivo"], data_root)
        if path is None:
            registrado_inexistente.append({"id": row["id"], "caminho_db": row["caminho_arquivo"]})
            continue
        registrado_existente += 1
        if path.name != row["nome_arquivo"]:
            nome_divergente.append({"id": row["id"], "db": row["nome_arquivo"], "fs": path.name})
        try:
            if row["hash_sha256"] and _file_sha256(path) != row["hash_sha256"]:
                hash_divergente.append({"id": row["id"], "caminho": str(path)})
        except OSError:
            registrado_inexistente.append({"id": row["id"], "caminho_db": row["caminho_arquivo"], "erro": "leitura"})

    db_paths = {str(_resolve_file_path(r["caminho_arquivo"], data_root)).lower() for r in rows}
    db_paths.discard("none")
    existente_sem_registro = [str(p) for p in fs_pdf if str(p).lower() not in db_paths]

    return {
        "registros_db": len(rows),
        "arquivos_fs": len(fs_pdf),
        "registrado_e_existente": registrado_existente,
        "registrado_inexistente": registrado_inexistente,
        "registrado_inexistente_total": len(registrado_inexistente),
        "existente_sem_registro": existente_sem_registro[:100],
        "existente_sem_registro_total": len(existente_sem_registro),
        "hash_divergente": hash_divergente,
        "hash_divergente_total": len(hash_divergente),
        "nome_divergente": nome_divergente,
    }


def _integrity_checks(conn: sqlite3.Connection) -> dict[str, Any]:
    checks: dict[str, int] = {}
    queries = {
        "fk_item_carregamento_orfa": """
            SELECT COUNT(*) FROM item_carregamento ic
            LEFT JOIN carregamento c ON c.id = ic.carregamento_id
            WHERE c.id IS NULL
        """,
        "fk_documento_carregamento_orfa": """
            SELECT COUNT(*) FROM documento d
            LEFT JOIN carregamento c ON c.id = d.carregamento_id
            WHERE c.id IS NULL
        """,
        "fk_item_nota_orfa": """
            SELECT COUNT(*) FROM item_nota_fiscal inf
            LEFT JOIN nota_fiscal nf ON nf.id = inf.nota_fiscal_id
            WHERE nf.id IS NULL
        """,
        "fk_item_carregamento_nf_orfa": """
            SELECT COUNT(*) FROM item_carregamento ic
            WHERE ic.nota_fiscal_id IS NOT NULL
              AND NOT EXISTS (SELECT 1 FROM nota_fiscal nf WHERE nf.id = ic.nota_fiscal_id)
        """,
        "carregamentos_sem_itens": """
            SELECT COUNT(*) FROM carregamento c
            WHERE NOT EXISTS (SELECT 1 FROM item_carregamento ic WHERE ic.carregamento_id = c.id)
        """,
        "notas_sem_itens": """
            SELECT COUNT(*) FROM nota_fiscal nf
            WHERE NOT EXISTS (SELECT 1 FROM item_nota_fiscal inf WHERE inf.nota_fiscal_id = nf.id)
        """,
        "historicos_orfaos": """
            SELECT COUNT(*) FROM historico_operacional h
            LEFT JOIN carregamento c ON c.id = h.carregamento_id
            WHERE c.id IS NULL
        """,
        "eventos_carregamento_orfaos": """
            SELECT COUNT(*) FROM evento_auditoria e
            WHERE e.entidade_tipo = 'carregamento'
              AND NOT EXISTS (SELECT 1 FROM carregamento c WHERE c.id = e.entidade_id)
        """,
        "xml_sem_chave": """
            SELECT COUNT(*) FROM documento_xml WHERE TRIM(chave_nfe) = ''
        """,
    }
    for name, sql in queries.items():
        try:
            checks[name] = int(conn.execute(sql).fetchone()[0])
        except sqlite3.Error:
            checks[name] = -1
    return checks


def _neon_prep(table_rows: list[dict[str, Any]], sqlite_bytes: int, total_rows: int) -> dict[str, Any]:
    batch = 500
    total_mb = sqlite_bytes / (1024 * 1024)
    est_seconds = max(total_rows / max(batch, 1) * 0.05, 1.0)
    return {
        "registros_por_tabela": {r["tabela"]: r["qtd_registros"] for r in table_rows},
        "total_registros": total_rows,
        "estimativa_crescimento_percentual": 15,
        "espaco_sqlite_mb": round(total_mb, 4),
        "espaco_previsto_neon_mb": round(total_mb * 1.35, 4),
        "batch_recomendado": batch,
        "tempo_etl_estimado_segundos": round(est_seconds, 2),
        "tempo_etl_estimado_minutos": round(est_seconds / 60, 2),
    }


def _build_risks(report: AuditReport) -> list[RiskItem]:
    risks: list[RiskItem] = []
    if not report.sqlite_found:
        risks.append(
            RiskItem(
                "CRITICO",
                "SQLITE_AUSENTE",
                "Banco SQLite de producao nao localizado neste ambiente.",
                "Migracao impossivel sem fonte de dados.",
            )
        )
    if report.fase_f_xml_consistencia.get("registrado_inexistente_total", 0) > 0:
        risks.append(
            RiskItem(
                "ALTO",
                "XML_DB_SEM_ARQUIVO",
                f"{report.fase_f_xml_consistencia['registrado_inexistente_total']} XML(s) registrados sem arquivo fisico.",
                "Perda aparente de consistencia banco x disco.",
            )
        )
    if report.fase_f_xml_consistencia.get("existente_sem_registro_total", 0) > 0:
        risks.append(
            RiskItem(
                "MEDIO",
                "XML_ARQUIVO_SEM_DB",
                f"{report.fase_f_xml_consistencia['existente_sem_registro_total']} XML(s) fisicos sem registro.",
                "Metadados ausentes apos migracao se nao reconciliados.",
            )
        )
    if report.fase_g_pdf_consistencia.get("registrado_inexistente_total", 0) > 0:
        risks.append(
            RiskItem(
                "ALTO",
                "PDF_DB_SEM_ARQUIVO",
                f"{report.fase_g_pdf_consistencia['registrado_inexistente_total']} PDF(s) registrados sem arquivo.",
                "Documentos inacessiveis na operacao.",
            )
        )
    for key, count in report.fase_i_integridade.items():
        if count > 0:
            risks.append(
                RiskItem(
                    "ALTO",
                    f"INTEGRIDADE_{key.upper()}",
                    f"{count} ocorrencia(s) em {key}.",
                    "Integridade referencial comprometida.",
                )
            )
    if report.fase_h_espaco.get("total_mb", 0) > 400:
        risks.append(
            RiskItem(
                "MEDIO",
                "CAPACIDADE_NEON",
                "Patrimonio proximo do limite operacional de 500 MB no Neon.",
                "Monitorar capacidade pos-migracao.",
            )
        )
    if not risks:
        risks.append(
            RiskItem("BAIXO", "SEM_RISCOS_MAPEADOS", "Nenhum risco adicional identificado na auditoria.", "Prosseguir com homologacao.")
        )
    return risks


def _render_markdown(report: AuditReport) -> str:
    lines = [
        "# Auditoria do Patrimonio de Dados (Pre-Migracao Neon)",
        "",
        f"**Gerado em:** {report.timestamp_utc}",
        f"**SQLite:** `{report.sqlite_path or 'NAO LOCALIZADO'}`",
        f"**Apto para migracao:** {'SIM' if report.apto_migracao else 'NAO'}",
        "",
        "## Resumo Executivo",
        "",
    ]
    for key, value in report.resumo.items():
        lines.append(f"- **{key}:** {value}")
    if report.bloqueador:
        lines.append(f"- **Bloqueador:** {report.bloqueador}")
    lines.extend(["", "## FASE A — Inventario de Tabelas", "", "| Tabela | Registros | Primeiro ID | Ultimo ID | MB |", "|--------|-----------|-------------|-----------|-----|"])
    for row in report.fase_a_tabelas:
        lines.append(
            f"| {row['tabela']} | {row['qtd_registros']} | {row.get('primeiro_id','—')} | {row.get('ultimo_id','—')} | {row.get('mb_aproximado',0)} |"
        )
    lines.extend(["", "## FASE B — documento_xml", ""])
    for key, value in report.fase_b_documento_xml.get("colunas_esperadas", {}).items():
        lines.append(f"- {key}: {'sim' if value else 'nao'}")
    lines.extend(["", "## FASE E — Diretorios fisicos", ""])
    for d in report.fase_e_diretorios:
        lines.append(f"- `{d['caminho']}` — {d['quantidade']} arquivos — {d['mb']} MB")
    lines.extend(["", "## FASE H — Espaco", ""])
    for key, value in report.fase_h_espaco.items():
        lines.append(f"- {key}: {value}")
    lines.extend(["", "## FASE I — Integridade", ""])
    for key, value in report.fase_i_integridade.items():
        lines.append(f"- {key}: {value}")
    lines.extend(["", "## FASE K — Riscos", ""])
    for risk in report.fase_k_riscos:
        lines.append(f"- **[{risk['nivel']}]** {risk['codigo']}: {risk['descricao']}")
    lines.extend(["", "## Conclusao", "", report.resumo.get("conclusao", "")])
    return "\n".join(lines)


def run_audit() -> AuditReport:
    start = time.perf_counter()
    report = AuditReport(timestamp_utc=datetime.now(timezone.utc).isoformat())
    db_path, attempts = _discover_sqlite_path()
    report.sqlite_path = str(db_path) if db_path else None
    report.sqlite_found = db_path is not None and db_path.is_file()

    data_root = (PROJECT_ROOT / "data").resolve()
    xml_dirs = []
    pdf_dirs = []
    for directory in _discover_asset_dirs():
        xml_scan = _scan_directory(directory, {".xml"})
        pdf_scan = _scan_directory(directory, {".pdf"})
        if xml_scan["quantidade"]:
            xml_dirs.append(xml_scan)
        if pdf_scan["quantidade"]:
            pdf_dirs.append(pdf_scan)
    report.fase_e_diretorios = sorted(
        {d["caminho"]: d for d in xml_dirs + pdf_dirs}.values(),
        key=lambda x: x["caminho"],
    )

    total_xml = sum(d["quantidade"] for d in xml_dirs)
    total_pdf = sum(d["quantidade"] for d in pdf_dirs)
    xml_bytes = sum(d["bytes"] for d in xml_dirs)
    pdf_bytes = sum(d["bytes"] for d in pdf_dirs)

    if not report.sqlite_found:
        report.bloqueador = (
            "Banco SQLite nao localizado. Tentativas: " + "; ".join(attempts[:6])
        )
        report.fase_h_espaco = {
            "sqlite_bytes": 0,
            "sqlite_mb": 0,
            "xml_bytes": xml_bytes,
            "xml_mb": round(xml_bytes / (1024 * 1024), 4),
            "pdf_bytes": pdf_bytes,
            "pdf_mb": round(pdf_bytes / (1024 * 1024), 4),
            "total_bytes": xml_bytes + pdf_bytes,
            "total_mb": round((xml_bytes + pdf_bytes) / (1024 * 1024), 4),
            "total_gb": round((xml_bytes + pdf_bytes) / (1024 * 1024 * 1024), 6),
        }
        report.fase_f_xml_consistencia = {
            "registros_db": 0,
            "arquivos_fs_chave": total_xml,
            "observacao": "Sem banco — comparacao DB x FS indisponivel.",
            "existente_sem_registro_total": total_xml,
        }
        report.fase_g_pdf_consistencia = {
            "registros_db": 0,
            "arquivos_fs": total_pdf,
            "observacao": "Sem banco — comparacao DB x FS indisponivel.",
            "existente_sem_registro_total": total_pdf,
        }
        report.apto_migracao = False
    else:
        conn = _open_sqlite_readonly(db_path)
        try:
            report.fase_a_tabelas = audit_tables(conn)
            total_rows = sum(r["qtd_registros"] for r in report.fase_a_tabelas)
            sqlite_bytes = db_path.stat().st_size

            if "documento_xml" in _table_names(conn):
                report.fase_b_documento_xml = _audit_table_detail(conn, "documento_xml")
            if "documento" in _table_names(conn):
                report.fase_c_documento = _audit_table_detail(conn, "documento")

            report.fase_d_notas = {}
            for table in ("nota_fiscal", "item_nota_fiscal", "item_carregamento", "carregamento"):
                if table in _table_names(conn):
                    report.fase_d_notas[table] = {
                        "quantidade": int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]),
                        "foreign_keys": _table_fks(conn, table),
                        "indices": _table_indexes(conn, table),
                    }

            data_root = db_path.parent.resolve()
            report.fase_f_xml_consistencia = _compare_xml(conn, data_root)
            report.fase_g_pdf_consistencia = _compare_pdf(conn, data_root)
            report.fase_i_integridade = _integrity_checks(conn)
            report.fase_j_neon_prep = _neon_prep(report.fase_a_tabelas, sqlite_bytes, total_rows)

            total_bytes = sqlite_bytes + xml_bytes + pdf_bytes
            report.fase_h_espaco = {
                "sqlite_bytes": sqlite_bytes,
                "sqlite_mb": round(sqlite_bytes / (1024 * 1024), 4),
                "xml_bytes": xml_bytes,
                "xml_mb": round(xml_bytes / (1024 * 1024), 4),
                "pdf_bytes": pdf_bytes,
                "pdf_mb": round(pdf_bytes / (1024 * 1024), 4),
                "total_bytes": total_bytes,
                "total_mb": round(total_bytes / (1024 * 1024), 4),
                "total_gb": round(total_bytes / (1024 * 1024 * 1024), 6),
                "percentual_sqlite": round((sqlite_bytes / total_bytes) * 100, 2) if total_bytes else 0,
                "percentual_xml": round((xml_bytes / total_bytes) * 100, 2) if total_bytes else 0,
                "percentual_pdf": round((pdf_bytes / total_bytes) * 100, 2) if total_bytes else 0,
            }

            crit_integridade = any(v > 0 for v in report.fase_i_integridade.values())
            crit_xml = report.fase_f_xml_consistencia.get("registrado_inexistente_total", 0) > 0
            crit_pdf = report.fase_g_pdf_consistencia.get("registrado_inexistente_total", 0) > 0
            report.apto_migracao = not (crit_integridade or crit_xml or crit_pdf)
            if not report.apto_migracao:
                report.bloqueador = "Inconsistencias criticas de integridade ou arquivos orfaos."
        finally:
            conn.close()

    risks = _build_risks(report)
    report.fase_k_riscos = [asdict(r) for r in risks]
    report.resumo = {
        "tabelas": len(report.fase_a_tabelas),
        "total_registros": sum(r["qtd_registros"] for r in report.fase_a_tabelas),
        "xml_fisicos": total_xml,
        "pdf_fisicos": total_pdf,
        "espaco_total_mb": report.fase_h_espaco.get("total_mb", 0),
        "problemas_criticos": sum(1 for r in risks if r.nivel == "CRITICO"),
        "problemas_altos": sum(1 for r in risks if r.nivel == "ALTO"),
        "apto_migracao": report.apto_migracao,
        "conclusao": (
            "Patrimonio apto para migracao ao PostgreSQL Neon."
            if report.apto_migracao
            else "Patrimonio NAO apto — resolver bloqueadores antes da migracao."
        ),
    }
    report.tempos_ms["total"] = round((time.perf_counter() - start) * 1000, 2)
    return report


def main() -> int:
    report = run_audit()
    REPORT_JSON.parent.mkdir(parents=True, exist_ok=True)
    REPORT_JSON.write_text(json.dumps(asdict(report), indent=2, ensure_ascii=False), encoding="utf-8")
    REPORT_MD.write_text(_render_markdown(report), encoding="utf-8")
    print(REPORT_MD.read_text(encoding="utf-8"))
    return 0 if report.sqlite_found else 2


if __name__ == "__main__":
    raise SystemExit(main())
