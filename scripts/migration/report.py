from __future__ import annotations

import json
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


@dataclass
class MigrationReport:
    fase: str = "N3"
    timestamp_utc: str = ""
    aprovada: bool = False
    bloqueador: str | None = None
    sqlite_source: dict[str, Any] = field(default_factory=dict)
    postgres_target: dict[str, Any] = field(default_factory=dict)
    inventario: dict[str, Any] = field(default_factory=dict)
    extracao: dict[str, Any] = field(default_factory=dict)
    validacao_pre_carga: dict[str, Any] = field(default_factory=dict)
    carga: dict[str, Any] = field(default_factory=dict)
    validacao_pos_carga: dict[str, Any] = field(default_factory=dict)
    arquivos: dict[str, Any] = field(default_factory=dict)
    tempos_ms: dict[str, float] = field(default_factory=dict)
    erros: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    rollback: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.timestamp_utc:
            self.timestamp_utc = datetime.now(timezone.utc).isoformat()

    def save(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)

        def _json_default(obj: object) -> str:
            return str(obj)

        path.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False, default=_json_default),
            encoding="utf-8",
        )
        return path


def compare_inventories(source: dict[str, Any], target: dict[str, Any]) -> dict[str, Any]:
    source_tables = {item["name"]: item for item in source.get("tables", [])}
    target_tables = {item["name"]: item for item in target.get("tables", [])}
    count_mismatches: list[str] = []
    checksum_warnings: list[str] = []
    for name, src in source_tables.items():
        tgt = target_tables.get(name, {})
        if src.get("row_count", 0) != tgt.get("row_count", 0):
            count_mismatches.append(
                f"count:{name}:sqlite={src.get('row_count')}:pg={tgt.get('row_count')}"
            )
        elif src.get("checksum") and src.get("checksum") != tgt.get("checksum"):
            checksum_warnings.append(f"checksum_repr:{name}")
    return {
        "equivalente": len(count_mismatches) == 0,
        "diferencas": count_mismatches,
        "avisos_checksum": checksum_warnings,
    }
