from __future__ import annotations

from pathlib import Path

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.schema import ensure_full_schema


def configure_application_storage(data_dir: Path) -> None:
    """Configura persistencia SQL da aplicacao."""
    settings = get_settings()
    configure_database(
        database_url=settings.database_url,
        echo=settings.echo_sql,
        data_root=data_dir,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=data_dir / "xml_storage",
    )
    ensure_full_schema()
    configure_auth_storage(data_dir)
    configure_carregamentos_storage(data_dir)
