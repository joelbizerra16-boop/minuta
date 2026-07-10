"""Testes da estabilizacao A1 — coerencia PostgreSQL como fonte oficial."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import pytest

from core.infrastructure_bootstrap import reset_infrastructure_bootstrap_state
from core.runtime_data_coherence import (
    get_classificacao_version_token,
    get_config_data_signature,
    get_reference_data_signature,
    get_xml_data_signature,
)
from core.settings import get_settings
from infrastructure.repositories.configuracao_repository import ConfiguracaoRecord
from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
from infrastructure.storage.config_storage import CONFIG_CHAVE_SEPARACAO, SqlJsonConfigStorage
from infrastructure.storage.xml_storage import SqlXmlRecordRepository


@pytest.fixture(autouse=True)
def _clean_bootstrap_state():
    get_settings.cache_clear()
    reset_infrastructure_bootstrap_state()
    yield
    reset_infrastructure_bootstrap_state()
    get_settings.cache_clear()


def test_xml_signature_changes_after_upsert() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "sig_xml.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        from core.bootstrap import configure_application_storage

        configure_application_storage()

        before = get_xml_data_signature()
        SqlXmlRecordRepository().upsert_records(
            [
                {
                    "NF": "123",
                    "ChaveNFe": "1" * 44,
                    "DataReferencia": "2026-01-01",
                    "DataReferenciaISO": "2026-01-01T00:00:00",
                    "Status": "autorizada",
                    "Items": [],
                }
            ]
        )
        after = get_xml_data_signature()

        assert after.count == before.count + 1
        assert after.revision is not None
        reset_infrastructure_bootstrap_state()


def test_config_signature_changes_after_save() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "sig_cfg.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        from core.bootstrap import configure_application_storage

        configure_application_storage()

        storage = SqlJsonConfigStorage()
        before = get_config_data_signature(CONFIG_CHAVE_SEPARACAO)
        storage.save_list(CONFIG_CHAVE_SEPARACAO, [{"NF": "1", "Chave": "2" * 22}])
        after = get_config_data_signature(CONFIG_CHAVE_SEPARACAO)

        assert after.count > before.count
        assert after.revision is not None
        reset_infrastructure_bootstrap_state()


def test_reference_signature_is_tuple_from_postgresql() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "sig_ref.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        from core.bootstrap import configure_application_storage

        configure_application_storage()

        signature = get_reference_data_signature()
        assert len(signature) == 2
        token = get_classificacao_version_token()
        assert isinstance(token[0], int)
        reset_infrastructure_bootstrap_state()


def test_config_record_revision_uses_atualizado_em() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        db_path = Path(tmp_dir) / "sig_rev.db"
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{db_path.as_posix()}"
        get_settings.cache_clear()

        from core.bootstrap import configure_application_storage

        configure_application_storage()

        repository = SqlConfiguracaoRepository()
        saved = repository.save(
            ConfiguracaoRecord(
                id=0,
                chave="teste.a1",
                valor='[{"ok": true}]',
                categoria="TESTE",
                tipo_valor="JSON",
            )
        )
        signature = get_config_data_signature("teste.a1")
        assert signature.revision is not None
        assert saved.atualizado_em is not None
        reset_infrastructure_bootstrap_state()
