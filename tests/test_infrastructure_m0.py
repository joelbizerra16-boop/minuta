#!/usr/bin/env python
"""Testes da infraestrutura M0 (sem alterar comportamento da aplicacao)."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from core.bootstrap import configure_application_storage
from core.settings import StorageBackend, get_settings
from infrastructure.database import configure_database, get_engine, get_session_factory
from infrastructure.models import Base
from infrastructure.unit_of_work import UnitOfWork


def test_default_storage_backend_is_sql() -> None:
    get_settings.cache_clear()
    settings = get_settings()
    assert settings.storage_backend == StorageBackend.SQL
    print("default storage backend SQL OK")


def test_bootstrap_configures_database() -> None:
    import infrastructure.database as db_module

    get_settings.cache_clear()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(tmp_path / 'boot.db').as_posix()}"
        get_settings.cache_clear()
        configure_application_storage()
        engine = get_engine()
        assert engine is not None
        engine.dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("bootstrap SQL OK")


def test_sqlalchemy_engine_and_metadata() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        data_root = Path(tmp_dir)
        db_path = data_root / "test.db"
        configure_database(
            database_url=f"sqlite:///{db_path.as_posix()}",
            data_root=data_root,
            pdf_storage_dir=data_root / "documentos",
        )
        engine = get_engine()
        Base.metadata.create_all(engine)

        factory = get_session_factory()
        assert factory is not None

        with UnitOfWork() as uow:
            assert uow.session is not None

        table_names = set(Base.metadata.tables.keys())
        expected = {
            "usuario",
            "perfil",
            "motorista",
            "veiculo",
            "destinatario",
            "rota",
            "nota_fiscal",
            "item_nota_fiscal",
            "carregamento",
            "item_carregamento",
            "documento",
            "historico_operacional",
            "evento_auditoria",
            "configuracao",
        }
        assert expected.issubset(table_names)

        engine.dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
    print("SQLAlchemy engine and ORM metadata OK")


def test_repository_contracts_importable() -> None:
    from infrastructure.repositories import (
        ConfiguracaoRepository,
        DocumentoRepository,
        HistoricoRepository,
        NotaFiscalRepository,
    )
    from infrastructure.repositories.sql import (
        SqlConfiguracaoRepository,
        SqlDocumentoRepository,
        SqlHistoricoRepository,
        SqlNotaFiscalRepository,
    )

    assert issubclass(SqlNotaFiscalRepository, NotaFiscalRepository)
    assert issubclass(SqlDocumentoRepository, DocumentoRepository)
    assert issubclass(SqlHistoricoRepository, HistoricoRepository)
    assert issubclass(SqlConfiguracaoRepository, ConfiguracaoRepository)
    print("repository contracts OK")


if __name__ == "__main__":
    test_default_storage_backend_is_sql()
    test_bootstrap_configures_database()
    test_sqlalchemy_engine_and_metadata()
    test_repository_contracts_importable()
    print("All M0 infrastructure tests passed.")
