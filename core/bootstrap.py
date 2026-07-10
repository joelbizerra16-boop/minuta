from __future__ import annotations

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.persistence.bootstrap_log import log_database_resolution
from infrastructure.schema import ensure_full_schema


def configure_application_storage() -> None:
    """Configura persistencia SQL da aplicacao a partir de MINUTA_DATABASE_URL."""
    settings = get_settings()
    log_database_resolution(
        runtime_environment=settings.runtime_environment.value,
        database_url_source=settings.database_url_source.value,
        database_url=settings.database_url,
    )
    configure_database(
        database_url=settings.database_url,
        echo=settings.echo_sql,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    ensure_full_schema()
    configure_auth_storage(settings.data_root)
    configure_carregamentos_storage(settings.data_root)
