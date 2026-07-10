"""Instrumentacao temporaria de operacoes SQL (ativar com MINUTA_SQL_AUDIT=1)."""

from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass, field
from threading import Lock
from typing import Any

from sqlalchemy import event
from sqlalchemy.engine import Engine

_LOGGER = logging.getLogger("minuta.persistence.sql_audit")

_ENABLED = os.getenv("MINUTA_SQL_AUDIT", "").strip().lower() in {"1", "true", "yes", "on"}
_LOCK = Lock()
_ENTRIES: list[dict[str, Any]] = []
_ACTIVE = False


@dataclass
class SqlAuditReport:
    entries: list[dict[str, Any]] = field(default_factory=list)

    def to_markdown(self) -> str:
        if not self.entries:
            return "Nenhuma operacao SQL registrada."

        lines = [
            "## Relatorio SQL Audit",
            "",
            "| # | Operacao | Tabela | Duracao (ms) | Origem |",
            "| --- | --- | --- | ---: | --- |",
        ]
        for index, entry in enumerate(self.entries, start=1):
            lines.append(
                f"| {index} | {entry.get('operation', '--')} | {entry.get('table', '--')} | "
                f"{float(entry.get('duration_ms', 0.0)):.2f} | {entry.get('origin', '--')} |"
            )
        return "\n".join(lines)


def is_sql_audit_enabled() -> bool:
    return _ENABLED


def reset_sql_audit() -> None:
    with _LOCK:
        _ENTRIES.clear()


def get_sql_audit_report() -> SqlAuditReport:
    with _LOCK:
        return SqlAuditReport(entries=list(_ENTRIES))


def _infer_operation(statement: str) -> str:
    token = str(statement or "").lstrip().split(None, 1)[0].upper() if statement else ""
    return token or "UNKNOWN"


def _infer_table(statement: str) -> str:
    text = " ".join(str(statement or "").split())
    upper = text.upper()
    for keyword in (" INTO ", " FROM ", " UPDATE ", " TABLE "):
        if keyword in upper:
            fragment = upper.split(keyword, 1)[1].strip()
            return fragment.split()[0].strip('"').strip("'")
    return "--"


def _record_entry(*, operation: str, table: str, duration_ms: float, origin: str) -> None:
    with _LOCK:
        _ENTRIES.append(
            {
                "operation": operation,
                "table": table,
                "duration_ms": duration_ms,
                "origin": origin,
            }
        )


def register_sql_audit(engine: Engine) -> None:
    global _ACTIVE
    if not _ENABLED or _ACTIVE:
        return

    @event.listens_for(engine, "before_cursor_execute")
    def _before_cursor_execute(
        conn: Any,
        cursor: Any,
        statement: str,
        parameters: Any,
        context: Any,
        executemany: bool,
    ) -> None:
        context._minuta_audit_started_at = time.perf_counter()

    @event.listens_for(engine, "after_cursor_execute")
    def _after_cursor_execute(
        conn: Any,
        cursor: Any,
        statement: str,
        parameters: Any,
        context: Any,
        executemany: bool,
    ) -> None:
        started_at = getattr(context, "_minuta_audit_started_at", None)
        duration_ms = 0.0
        if started_at is not None:
            duration_ms = (time.perf_counter() - started_at) * 1000.0
        origin = str(getattr(context, "executemany", executemany))
        _record_entry(
            operation=_infer_operation(statement),
            table=_infer_table(statement),
            duration_ms=duration_ms,
            origin=origin,
        )

    _ACTIVE = True
    _LOGGER.info("sql_audit enabled for engine dialect=%s", engine.dialect.name)
