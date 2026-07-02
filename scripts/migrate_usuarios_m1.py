#!/usr/bin/env python
"""Script de migracao controlada dos usuarios (Fase M1)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from auth.migration.migrate_usuarios import migrate_usuarios_from_json
from core.settings import get_settings


def main() -> int:
    parser = argparse.ArgumentParser(description="Migracao M1 de usuarios JSON -> SQL")
    parser.add_argument(
        "--json-path",
        type=Path,
        default=BASE_DIR / "data" / "usuarios.json",
        help="Caminho do usuarios.json",
    )
    parser.add_argument(
        "--skip-backup",
        action="store_true",
        help="Nao criar backup do JSON (nao recomendado)",
    )
    args = parser.parse_args()

    settings = get_settings()
    report = migrate_usuarios_from_json(
        args.json_path,
        database_url=settings.database_url,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        create_backup=not args.skip_backup,
    )
    print(report.to_text())
    return 0 if report.success else 1


if __name__ == "__main__":
    raise SystemExit(main())
