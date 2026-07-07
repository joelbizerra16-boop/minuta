from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_carregamento_service, get_rastreabilidade_nf_service
from carregamentos.repository.sql_rastreabilidade_nf_repository import (
    SqlRastreabilidadeNfRepository,
    _build_rastreabilidade_nf_sql,
)
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'carregamentos.db').as_posix()}"
    get_settings.cache_clear()
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=data_dir,
        pdf_storage_dir=data_dir / "documentos",
        xml_storage_dir=data_dir / "xml_storage",
    )
    ensure_full_schema()
    configure_auth_storage(data_dir)
    configure_carregamentos_storage(data_dir)


def _teardown_sql_env() -> None:
    import carregamentos.bootstrap as carregamentos_bootstrap
    import infrastructure.database as db_module

    carregamentos_bootstrap.get_carregamento_service.cache_clear()
    carregamentos_bootstrap.get_rastreabilidade_nf_service.cache_clear()
    carregamentos_bootstrap._repository = None
    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None


def _build_processed_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "NF": "1001",
                "cProd": "P001",
                "Descricao": "Oleo A",
                "Qtd": 2,
                "Unidade": "UN",
                "Peso": 500.0,
                "Destinatario": "Cliente A",
                "ROTA": "VALE BRIDA",
                "ChaveNFe": "35260600000000000000000000000000000000000001",
                "Status": "Autorizado o uso da NF-e",
            },
        ]
    )


def _build_summary(numero_carga: str = "000154") -> dict:
    return {
        "numero_carga": numero_carga,
        "motorista": "Carlos Silva",
        "placa": "ABC-1234",
        "filial": "BRIDA",
        "data_saida": "29/06/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 500.0,
    }


def _run_test(fn) -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        try:
            _setup_sql_env(data_dir)
            fn(data_dir)
        finally:
            _teardown_sql_env()


def test_rastreabilidade_sql_agrega_carregamentos_distintos() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()

        primeiro = service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=b"%PDF-romaneio%",
        )
        assert primeiro is not None

        service.register_from_processing(
            summary={**summary, "numero_carga": "000155"},
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
            is_reentrega=True,
        )

        relatorio = SqlRastreabilidadeNfRepository().buscar_por_termo("1001")
        assert relatorio is not None
        assert relatorio.resumo.numero_nf == "1001"
        assert relatorio.resumo.quantidade_carregamentos == 2
        assert relatorio.resumo.quantidade_reentregas == 1
        assert len(relatorio.historico) == 2
        assert relatorio.estatisticas is not None
        assert relatorio.estatisticas.total_reentregas == 1
        assert any(item.reentrega for item in relatorio.historico)
        assert "json_group_array" in _build_rastreabilidade_nf_sql("sqlite")
        print("rastreabilidade sql OK")

    _run_test(_case)


def test_rastreabilidade_pdf_bytes() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        saved = service.register_from_processing(
            summary=_build_summary(),
            processed_df=_build_processed_df(),
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=b"%PDF-romaneio%",
        )
        assert saved is not None

        pdf_bytes = get_rastreabilidade_nf_service().gerar_relatorio_pdf("1001", current_user=None)
        assert pdf_bytes.startswith(b"%PDF")
        assert len(pdf_bytes) > 500
        print("rastreabilidade pdf OK")

    _run_test(_case)


def test_rastreabilidade_termo_inexistente() -> None:
    def _case(_: Path) -> None:
        relatorio = SqlRastreabilidadeNfRepository().buscar_por_termo("999999")
        assert relatorio is None
        print("rastreabilidade vazia OK")

    _run_test(_case)


if __name__ == "__main__":
    test_rastreabilidade_sql_agrega_carregamentos_distintos()
    test_rastreabilidade_pdf_bytes()
    test_rastreabilidade_termo_inexistente()
