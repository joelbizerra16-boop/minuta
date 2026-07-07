#!/usr/bin/env python
"""
Homologacao final do banco de producao (SOMENTE LEITURA).

Banco oficial: C:\\MinutaData\\minuta_dev.db
XML: ProjetoBrida\\01_Minuta\\data\\xml_storage
PDF: C:\\MinutaData\\documentos

Gera:
  reports/auditoria_patrimonio_dados.json
  reports/auditoria_patrimonio_dados.md
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OFFICIAL_DB = Path(r"C:\MinutaData\minuta_dev.db")
DATA_ROOT = Path(r"C:\MinutaData")
XML_DIR = PROJECT_ROOT / "data" / "xml_storage"
PDF_DIR = DATA_ROOT / "documentos"
PATH_BASES = [DATA_ROOT, DATA_ROOT / "documentos", PROJECT_ROOT / "data", PROJECT_ROOT]
REPORT_JSON = PROJECT_ROOT / "reports" / "auditoria_patrimonio_dados.json"
REPORT_MD = PROJECT_ROOT / "reports" / "auditoria_patrimonio_dados.md"

DOMAIN_TABLES = (
    "perfil", "usuario", "configuracao", "motorista", "veiculo", "destinatario", "rota",
    "nota_fiscal", "item_nota_fiscal", "carregamento", "item_carregamento",
    "documento", "documento_xml", "historico_operacional", "evento_auditoria",
)


@dataclass
class Report:
    timestamp_utc: str = ""
    banco_oficial: str = str(OFFICIAL_DB)
    banco_localizado: bool = False
    banco_homologado: bool = False
    apto_migracao: bool = False
    bloqueador: str | None = None
    inventario_tabelas: list[dict[str, Any]] = field(default_factory=list)
    documento_xml: dict[str, Any] = field(default_factory=dict)
    documento_pdf: dict[str, Any] = field(default_factory=dict)
    dominio: dict[str, Any] = field(default_factory=dict)
    integridade: dict[str, Any] = field(default_factory=dict)
    espaco: dict[str, Any] = field(default_factory=dict)
    riscos: list[dict[str, str]] = field(default_factory=list)
    resumo: dict[str, Any] = field(default_factory=dict)
    tempos_ms: dict[str, float] = field(default_factory=dict)


def _connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{db_path.as_posix()}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA query_only=ON")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _table_info(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    return [
        {
            "name": r[1],
            "type": r[2],
            "notnull": bool(r[3]),
            "pk": bool(r[5]),
        }
        for r in conn.execute(f"PRAGMA table_info([{table}])").fetchall()
    ]


def _fks(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    return [
        {"from": r[3], "to_table": r[2], "to_col": r[4], "on_delete": r[6]}
        for r in conn.execute(f"PRAGMA foreign_key_list([{table}])").fetchall()
    ]


def _indexes(conn: sqlite3.Connection, table: str) -> list[dict[str, Any]]:
    result = []
    for row in conn.execute(f"PRAGMA index_list([{table}])").fetchall():
        cols = [r[2] for r in conn.execute(f"PRAGMA index_info([{row[1]}])").fetchall()]
        result.append({"name": row[1], "unique": bool(row[2]), "columns": cols})
    return result


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _resolve_path(stored: str, bases: list[Path]) -> Path | None:
    raw = str(stored or "").strip()
    if not raw:
        return None
    p = Path(raw).expanduser()
    if p.is_file():
        return p.resolve()
    for base in bases:
        c = (base / raw).resolve()
        if c.is_file():
            return c
    return None


def inventory(conn: sqlite3.Connection, db_path: Path) -> list[dict[str, Any]]:
    tables = [
        r[0]
        for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        ).fetchall()
    ]
    db_bytes = db_path.stat().st_size
    rows_out = []
    for table in tables:
        count = int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0])
        info = _table_info(conn, table)
        pk = [c["name"] for c in info if c["pk"]]
        first_id = last_id = None
        if pk:
            first_id = conn.execute(f'SELECT MIN("{pk[0]}") FROM "{table}"').fetchone()[0]
            last_id = conn.execute(f'SELECT MAX("{pk[0]}") FROM "{table}"').fetchone()[0]
        est_mb = round((count * 300) / (1024 * 1024), 4) if count else 0.0
        rows_out.append(
            {
                "tabela": table,
                "qtd_registros": count,
                "primeiro_id": first_id,
                "ultimo_id": last_id,
                "mb_estimado": est_mb,
                "primary_key": pk,
                "foreign_keys": _fks(conn, table),
                "indices": _indexes(conn, table),
            }
        )
    return rows_out


def audit_documento_xml(conn: sqlite3.Connection) -> dict[str, Any]:
    info = _table_info(conn, "documento_xml")
    count = int(conn.execute("SELECT COUNT(*) FROM documento_xml").fetchone()[0])
    samples = [dict(r) for r in conn.execute("SELECT * FROM documento_xml ORDER BY id LIMIT 5").fetchall()]

    rows = conn.execute(
        "SELECT id, chave_nfe, nome_arquivo, caminho_arquivo, hash_sha256, tamanho, ativo FROM documento_xml"
    ).fetchall()

    fs_xml: dict[str, Path] = {}
    if XML_DIR.is_dir():
        for p in XML_DIR.rglob("*.xml"):
            key = p.stem
            if len(key) == 44:
                fs_xml[key] = p.resolve()

    bases = PATH_BASES
    ok, missing, hash_bad, path_bad = [], [], [], []
    for row in rows:
        chave = str(row["chave_nfe"])
        path = _resolve_path(row["caminho_arquivo"], bases)
        if path is None or not path.is_file():
            alt = fs_xml.get(chave)
            if alt and alt.is_file():
                ok.append({"chave_nfe": chave, "via": "xml_storage", "path": str(alt)})
            else:
                missing.append({"id": row["id"], "chave_nfe": chave, "caminho_db": row["caminho_arquivo"]})
            continue
        ok.append({"chave_nfe": chave, "path": str(path)})
        if row["nome_arquivo"] and path.name != row["nome_arquivo"]:
            path_bad.append({"chave_nfe": chave, "db": row["nome_arquivo"], "fs": path.name})
        try:
            if row["hash_sha256"] and _sha256(path) != row["hash_sha256"]:
                hash_bad.append({"chave_nfe": chave, "path": str(path)})
        except OSError:
            missing.append({"chave_nfe": chave, "erro": "leitura"})

    fs_only = [{"chave_nfe": k, "path": str(v)} for k, v in fs_xml.items() if k not in {str(r["chave_nfe"]) for r in rows}]

    return {
        "quantidade": count,
        "colunas": info,
        "colunas_esperadas": {
            "caminho_arquivo": True,
            "hash_sha256": True,
            "tamanho": True,
            "chave_nfe": True,
            "data_importacao": any(c["name"] == "data_importacao" for c in info),
            "xml_blob": any("blob" in c["type"].lower() for c in info),
        },
        "exemplos": samples,
        "cruzamento_xml_storage": {
            "diretorio": str(XML_DIR),
            "arquivos_fs": len(fs_xml),
            "registrado_e_existente": len(ok),
            "registrado_inexistente": missing[:50],
            "registrado_inexistente_total": len(missing),
            "fs_sem_registro": fs_only[:50],
            "fs_sem_registro_total": len(fs_only),
            "hash_divergente": hash_bad[:50],
            "hash_divergente_total": len(hash_bad),
            "nome_divergente": path_bad[:50],
            "nome_divergente_total": len(path_bad),
        },
    }


def audit_documento(conn: sqlite3.Connection) -> dict[str, Any]:
    info = _table_info(conn, "documento")
    count = int(conn.execute("SELECT COUNT(*) FROM documento").fetchone()[0])
    samples = [dict(r) for r in conn.execute("SELECT * FROM documento ORDER BY id LIMIT 5").fetchall()]
    rows = conn.execute(
        "SELECT id, carregamento_id, tipo, nome_arquivo, caminho_arquivo, hash_sha256 FROM documento"
    ).fetchall()

    fs_pdf = [p.resolve() for p in PDF_DIR.rglob("*.pdf")] if PDF_DIR.is_dir() else []
    bases = PATH_BASES

    ok, missing, hash_bad, path_bad = 0, [], [], []
    db_paths: set[str] = set()
    for row in rows:
        path = _resolve_path(row["caminho_arquivo"], bases)
        if path is None or not path.is_file():
            missing.append({"id": row["id"], "caminho_db": row["caminho_arquivo"]})
            continue
        ok += 1
        db_paths.add(str(path).lower())
        if row["nome_arquivo"] and path.name != row["nome_arquivo"]:
            path_bad.append({"id": row["id"], "db": row["nome_arquivo"], "fs": path.name})
        try:
            if row["hash_sha256"] and _sha256(path) != row["hash_sha256"]:
                hash_bad.append({"id": row["id"], "path": str(path)})
        except OSError:
            missing.append({"id": row["id"], "erro": "leitura"})

    fs_only = [str(p) for p in fs_pdf if str(p).lower() not in db_paths]

    return {
        "quantidade": count,
        "colunas": info,
        "exemplos": samples,
        "cruzamento_documentos": {
            "diretorio": str(PDF_DIR),
            "arquivos_fs": len(fs_pdf),
            "registrado_e_existente": ok,
            "registrado_inexistente": missing,
            "registrado_inexistente_total": len(missing),
            "fs_sem_registro": fs_only[:50],
            "fs_sem_registro_total": len(fs_only),
            "hash_divergente": hash_bad,
            "hash_divergente_total": len(hash_bad),
            "nome_divergente": path_bad,
        },
    }


def audit_domain(conn: sqlite3.Connection) -> dict[str, Any]:
    tables = [
        "carregamento", "item_carregamento", "nota_fiscal", "item_nota_fiscal",
        "historico_operacional", "evento_auditoria", "configuracao",
        "usuario", "motorista", "veiculo", "rota", "perfil",
    ]
    out = {}
    for table in tables:
        try:
            out[table] = {
                "quantidade": int(conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]),
                "foreign_keys": _fks(conn, table),
                "indices": _indexes(conn, table),
            }
        except sqlite3.Error:
            out[table] = {"quantidade": -1, "erro": "tabela ausente"}
    return out


def audit_integrity(conn: sqlite3.Connection) -> dict[str, Any]:
    queries = {
        "fk_item_carregamento_orfa": """
            SELECT COUNT(*) FROM item_carregamento ic
            LEFT JOIN carregamento c ON c.id = ic.carregamento_id WHERE c.id IS NULL""",
        "fk_documento_carregamento_orfa": """
            SELECT COUNT(*) FROM documento d
            LEFT JOIN carregamento c ON c.id = d.carregamento_id WHERE c.id IS NULL""",
        "fk_item_nota_orfa": """
            SELECT COUNT(*) FROM item_nota_fiscal inf
            LEFT JOIN nota_fiscal nf ON nf.id = inf.nota_fiscal_id WHERE nf.id IS NULL""",
        "carregamentos_sem_itens": """
            SELECT COUNT(*) FROM carregamento c
            WHERE NOT EXISTS (SELECT 1 FROM item_carregamento ic WHERE ic.carregamento_id = c.id)""",
        "notas_sem_itens": """
            SELECT COUNT(*) FROM nota_fiscal nf
            WHERE NOT EXISTS (SELECT 1 FROM item_nota_fiscal inf WHERE inf.nota_fiscal_id = nf.id)""",
        "historicos_orfaos": """
            SELECT COUNT(*) FROM historico_operacional h
            LEFT JOIN carregamento c ON c.id = h.carregamento_id WHERE c.id IS NULL""",
        "eventos_carregamento_orfaos": """
            SELECT COUNT(*) FROM evento_auditoria e
            WHERE e.entidade_tipo='carregamento'
              AND NOT EXISTS (SELECT 1 FROM carregamento c WHERE c.id=e.entidade_id)""",
        "itens_sem_carregamento": """
            SELECT COUNT(*) FROM item_carregamento ic WHERE ic.carregamento_id IS NULL""",
    }
    result = {}
    for name, sql in queries.items():
        result[name] = int(conn.execute(sql).fetchone()[0])
    result["fk_check"] = conn.execute("PRAGMA foreign_key_check").fetchall()
    return result


def run() -> Report:
    t0 = time.perf_counter()
    report = Report(timestamp_utc=datetime.now(timezone.utc).isoformat())

    if not OFFICIAL_DB.is_file():
        report.bloqueador = f"Banco oficial nao encontrado: {OFFICIAL_DB}"
        report.riscos.append({"nivel": "CRITICO", "codigo": "DB_AUSENTE", "descricao": report.bloqueador})
        return report

    report.banco_localizado = True
    conn = _connect(OFFICIAL_DB)
    try:
        report.inventario_tabelas = inventory(conn, OFFICIAL_DB)
        report.documento_xml = audit_documento_xml(conn)
        report.documento_pdf = audit_documento(conn)
        report.dominio = audit_domain(conn)
        report.integridade = audit_integrity(conn)

        db_bytes = OFFICIAL_DB.stat().st_size
        xml_bytes = sum(p.stat().st_size for p in XML_DIR.rglob("*.xml")) if XML_DIR.is_dir() else 0
        pdf_bytes = sum(p.stat().st_size for p in PDF_DIR.rglob("*.pdf")) if PDF_DIR.is_dir() else 0
        total = db_bytes + xml_bytes + pdf_bytes
        report.espaco = {
            "sqlite_bytes": db_bytes,
            "sqlite_mb": round(db_bytes / (1024 * 1024), 4),
            "xml_bytes": xml_bytes,
            "xml_mb": round(xml_bytes / (1024 * 1024), 4),
            "pdf_bytes": pdf_bytes,
            "pdf_mb": round(pdf_bytes / (1024 * 1024), 4),
            "total_mb": round(total / (1024 * 1024), 4),
            "total_gb": round(total / (1024 * 1024 * 1024), 6),
        }

        orphans = sum(
            v for k, v in report.integridade.items()
            if isinstance(v, int) and v > 0 and k != "notas_sem_itens"
        )
        xml_miss = report.documento_xml["cruzamento_xml_storage"]["registrado_inexistente_total"]
        pdf_miss = report.documento_pdf["cruzamento_documentos"]["registrado_inexistente_total"]
        hash_xml = report.documento_xml["cruzamento_xml_storage"]["hash_divergente_total"]
        hash_pdf = report.documento_pdf["cruzamento_documentos"]["hash_divergente_total"]
        fk_violations = len(report.integridade.get("fk_check", []))

        hash_xml_real = [
            x for x in report.documento_xml["cruzamento_xml_storage"].get("hash_divergente", [])
            if not str(x.get("chave_nfe", "")).startswith("1111111111111111111111111111111111111111111")
        ]
        hash_xml_bench = report.documento_xml["cruzamento_xml_storage"]["hash_divergente_total"] - len(hash_xml_real)

        report.banco_homologado = True
        report.apto_migracao = (
            orphans == 0
            and xml_miss == 0
            and pdf_miss == 0
            and len(hash_xml_real) == 0
            and hash_pdf == 0
            and fk_violations == 0
        )

        if not report.apto_migracao:
            parts = []
            if orphans:
                parts.append(f"integridade_fk_orfa={orphans}")
            if xml_miss:
                parts.append(f"xml_ausente={xml_miss}")
            if pdf_miss:
                parts.append(f"pdf_ausente={pdf_miss}")
            if hash_xml_real:
                parts.append(f"hash_xml_real={len(hash_xml_real)}")
            if hash_pdf:
                parts.append(f"hash_pdf={hash_pdf}")
            if fk_violations:
                parts.append(f"fk_check={fk_violations}")
            report.bloqueador = "Inconsistencias: " + ", ".join(parts) if parts else None

        if hash_xml_bench:
            report.riscos.append(
                {
                    "nivel": "BAIXO",
                    "codigo": "XML_BENCH_HASH",
                    "descricao": f"{hash_xml_bench} XML(s) de benchmark com hash sintetico (bench-N).",
                }
            )
        if report.integridade.get("notas_sem_itens", 0) > 0:
            report.riscos.append(
                {
                    "nivel": "MEDIO",
                    "codigo": "NOTAS_SEM_ITENS",
                    "descricao": f"{report.integridade['notas_sem_itens']} nota(s)_fiscal sem item_nota_fiscal.",
                }
            )

        total_rows = sum(t["qtd_registros"] for t in report.inventario_tabelas)
        report.resumo = {
            "banco_oficial": str(OFFICIAL_DB),
            "banco_localizado": True,
            "banco_homologado": report.banco_homologado,
            "apto_migracao": report.apto_migracao,
            "tabelas": len(report.inventario_tabelas),
            "total_registros": total_rows,
            "documento_xml": report.documento_xml.get("quantidade", 0),
            "documento_pdf": report.documento_pdf.get("quantidade", 0),
            "xml_fisicos": report.documento_xml.get("cruzamento_xml_storage", {}).get("arquivos_fs", 0),
            "pdf_fisicos": report.documento_pdf.get("cruzamento_documentos", {}).get("arquivos_fs", 0),
            "espaco_total_mb": report.espaco.get("total_mb", 0),
            "conclusao": (
                "Banco oficial localizado e homologado. Apto para migracao PostgreSQL Neon."
                if report.apto_migracao
                else "Banco oficial localizado e homologado com ressalvas. Revisar inconsistencias antes da migracao."
            ),
        }

        if xml_miss:
            report.riscos.append({"nivel": "ALTO", "codigo": "XML_AUSENTE", "descricao": f"{xml_miss} XML(s) no DB sem arquivo."})
        if pdf_miss:
            report.riscos.append({"nivel": "ALTO", "codigo": "PDF_AUSENTE", "descricao": f"{pdf_miss} PDF(s) no DB sem arquivo."})
        if hash_xml or hash_pdf:
            report.riscos.append({"nivel": "ALTO", "codigo": "HASH_DIVERGENTE", "descricao": f"hash_xml={hash_xml}, hash_pdf={hash_pdf}"})
        if orphans or fk_violations:
            report.riscos.append({"nivel": "ALTO", "codigo": "INTEGRIDADE", "descricao": f"orfaos={orphans}, fk_check={fk_violations}"})
        if report.apto_migracao:
            report.riscos.append({"nivel": "BAIXO", "codigo": "OK", "descricao": "Nenhuma inconsistencia critica detectada."})
    finally:
        conn.close()

    report.tempos_ms["total"] = round((time.perf_counter() - t0) * 1000, 2)
    return report


def _md(r: Report) -> str:
    lines = [
        "# Homologacao do Banco de Producao",
        "",
        f"**Banco oficial:** `{r.banco_oficial}`",
        f"**Localizado:** {'SIM' if r.banco_localizado else 'NAO'}",
        f"**Homologado:** {'SIM' if r.banco_homologado else 'NAO'}",
        f"**Apto migracao:** {'SIM' if r.apto_migracao else 'NAO'}",
        f"**Gerado em:** {r.timestamp_utc}",
        "",
        "## Resumo",
    ]
    for k, v in r.resumo.items():
        lines.append(f"- **{k}:** {v}")
    if r.bloqueador:
        lines.append(f"- **bloqueador:** {r.bloqueador}")
    lines.extend(["", "## Inventario de Tabelas", "", "| Tabela | Registros | Primeiro ID | Ultimo ID | MB est. |", "|--------|-----------|-------------|-----------|---------|"])
    for row in r.inventario_tabelas:
        lines.append(f"| {row['tabela']} | {row['qtd_registros']} | {row.get('primeiro_id','—')} | {row.get('ultimo_id','—')} | {row.get('mb_estimado',0)} |")
    cx = r.documento_xml.get("cruzamento_xml_storage", {})
    lines.extend(["", "## documento_xml x xml_storage", "", f"- Registros DB: {r.documento_xml.get('quantidade',0)}", f"- Arquivos FS: {cx.get('arquivos_fs',0)}", f"- OK: {cx.get('registrado_e_existente',0)}", f"- Ausentes: {cx.get('registrado_inexistente_total',0)}", f"- FS sem registro: {cx.get('fs_sem_registro_total',0)}", f"- Hash divergente: {cx.get('hash_divergente_total',0)}"])
    cp = r.documento_pdf.get("cruzamento_documentos", {})
    lines.extend(["", "## documento x documentos", "", f"- Registros DB: {r.documento_pdf.get('quantidade',0)}", f"- Arquivos FS: {cp.get('arquivos_fs',0)}", f"- OK: {cp.get('registrado_e_existente',0)}", f"- Ausentes: {cp.get('registrado_inexistente_total',0)}", f"- Hash divergente: {cp.get('hash_divergente_total',0)}"])
    lines.extend(["", "## Integridade", ""])
    for k, v in r.integridade.items():
        if k != "fk_check":
            lines.append(f"- {k}: {v}")
    lines.extend(["", "## Riscos", ""])
    for risk in r.riscos:
        lines.append(f"- **[{risk['nivel']}]** {risk['codigo']}: {risk['descricao']}")
    lines.extend(["", "## Conclusao", "", r.resumo.get("conclusao", "")])
    return "\n".join(lines)


def main() -> int:
    report = run()
    REPORT_JSON.parent.mkdir(parents=True, exist_ok=True)
    REPORT_JSON.write_text(json.dumps(asdict(report), indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    REPORT_MD.write_text(_md(report), encoding="utf-8")
    print(_md(report))
    return 0 if report.banco_homologado else 1


if __name__ == "__main__":
    raise SystemExit(main())
