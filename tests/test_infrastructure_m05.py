#!/usr/bin/env python
"""Validacao estrutural da Fase M0.5 — modelagem consolidada."""

from __future__ import annotations

import tempfile
from pathlib import Path

from sqlalchemy import inspect

from infrastructure.database import configure_database, get_engine
from infrastructure.models import Base
from infrastructure.models.constants import (
    AUDIT_CATEGORIA_AUTH,
    AUDIT_EVENTO_LOGIN,
    DOC_TIPO_MINUTA,
    ON_DELETE_RESTRICT,
)


def test_schema_tables_and_fk_policies() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        data_root = Path(tmp_dir)
        db_path = data_root / "m05.db"
        configure_database(
            database_url=f"sqlite:///{db_path.as_posix()}",
            data_root=data_root,
            pdf_storage_dir=data_root / "documentos",
            xml_storage_dir=data_root / "xml_storage",
        )
        engine = get_engine()
        Base.metadata.create_all(engine)

        table_names = set(Base.metadata.tables.keys())
        expected_tables = {
            "perfil",
            "usuario",
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
            "documento_xml",
        }
        assert expected_tables == table_names

        inspector = inspect(engine)

        usuario_fks = {fk["constrained_columns"][0]: fk for fk in inspector.get_foreign_keys("usuario")}
        assert usuario_fks["perfil_id"]["options"].get("ondelete") == ON_DELETE_RESTRICT.upper()

        carregamento_uq = {c["name"] for c in inspector.get_unique_constraints("carregamento")}
        assert "uq_carregamento_numero" in carregamento_uq

        nf_uq = {c["name"] for c in inspector.get_unique_constraints("nota_fiscal")}
        assert "uq_nota_fiscal_chave_nfe" in nf_uq

        doc_uq = {c["name"] for c in inspector.get_unique_constraints("documento")}
        assert "uq_documento_carregamento_tipo" in doc_uq

        usuario_cols = {col["name"] for col in inspector.get_columns("usuario")}
        assert {"ativo", "excluido_em", "perfil_id"}.issubset(usuario_cols)

        motorista_cols = {col["name"] for col in inspector.get_columns("motorista")}
        assert {"ativo", "excluido_em"}.issubset(motorista_cols)

        evento_cols = {col["name"] for col in inspector.get_columns("evento_auditoria")}
        assert {"categoria", "evento", "metadados_json", "entidade_tipo", "entidade_id"}.issubset(evento_cols)

        config_cols = {col["name"] for col in inspector.get_columns("configuracao")}
        assert {"categoria", "tipo_valor", "descricao", "criado_em", "atualizado_em"}.issubset(config_cols)

        engine.dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None

    print("schema tables and FK policies OK")


def test_audit_and_document_constants() -> None:
    assert AUDIT_CATEGORIA_AUTH == "AUTH"
    assert AUDIT_EVENTO_LOGIN == "LOGIN"
    assert DOC_TIPO_MINUTA == "MINUTA"
    print("domain constants OK")


def test_evento_auditoria_repository_contract() -> None:
    from infrastructure.repositories import EventoAuditoriaRepository

    assert EventoAuditoriaRepository.__abstractmethods__ == frozenset(
        {"append", "list_by_entidade", "list_by_usuario"}
    )
    print("evento auditoria contract OK")


if __name__ == "__main__":
    test_schema_tables_and_fk_policies()
    test_audit_and_document_constants()
    test_evento_auditoria_repository_contract()
    print("All M0.5 infrastructure tests passed.")
