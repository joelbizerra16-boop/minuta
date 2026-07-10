from __future__ import annotations

import logging

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage
from core.infrastructure_bootstrap import (
    get_bootstrapped_database_url,
    is_app_bootstrapped,
    mark_app_bootstrapped,
    reset_app_bootstrap_state,
)
from core.settings import get_settings
from infrastructure.database import configure_database
from infrastructure.persistence.bootstrap_log import log_database_resolution
from infrastructure.schema import ensure_full_schema

_LOGGER = logging.getLogger("minuta.bootstrap")


def reset_application_bootstrap_state() -> None:
    """Utilitario de teste: permite reexecutar configure_application_storage."""
    reset_app_bootstrap_state()


def configure_application_storage() -> None:
    """Configura persistencia SQL da aplicacao a partir de MINUTA_DATABASE_URL (once-per-process)."""
    settings = get_settings()
    database_url = settings.database_url

    if is_app_bootstrapped() and get_bootstrapped_database_url() == database_url:
        _LOGGER.debug("bootstrap.configure_application_storage skipped database_url=%s", database_url)
        return

    if is_app_bootstrapped() and get_bootstrapped_database_url() != database_url:
        from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state

        _LOGGER.info(
            "bootstrap.configure_application_storage database_url_changed old=%s new=%s",
            get_bootstrapped_database_url(),
            database_url,
        )
        reset_infrastructure_bootstrap_state()

    log_database_resolution(
        runtime_environment=settings.runtime_environment.value,
        database_url_source=settings.database_url_source.value,
        database_url=database_url,
    )
    configure_database(
        database_url=database_url,
        echo=settings.echo_sql,
        data_root=settings.data_root,
        pdf_storage_dir=settings.pdf_storage_dir,
        xml_storage_dir=settings.xml_storage_dir,
    )
    ensure_full_schema()
    configure_auth_storage(settings.data_root)
    configure_carregamentos_storage(settings.data_root)
    mark_app_bootstrapped(database_url=database_url)
    _LOGGER.info("bootstrap.configure_application_storage complete database_url=%s", database_url)
