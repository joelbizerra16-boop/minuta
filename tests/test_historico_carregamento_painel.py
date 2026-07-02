from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_fechamento_service, get_historico_carregamento_service
from carregamentos.models.operacional import DecisaoOperacional
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'historico_painel.db').as_posix()}"
    get_settings.cache_clear()
    configure_database(
        database_url=os.environ["MINUTA_DATABASE_URL"],
        data_root=data_dir,
        pdf_storage_dir=data_dir / "documentos",
    )
    ensure_full_schema()
    configure_auth_storage(data_dir)
    configure_carregamentos_storage(data_dir)


def _teardown_sql_env() -> None:
    import carregamentos.bootstrap as carregamentos_bootstrap
    import infrastructure.database as db_module

    carregamentos_bootstrap.get_carregamento_service.cache_clear()
    carregamentos_bootstrap._repository = None
    carregamentos_bootstrap._fechamento_service = None
    carregamentos_bootstrap._analise_operacional_service = None
    carregamentos_bootstrap._historico_carregamento_service = None
    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None


def _build_processed_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "NF": "4001",
                "cProd": "P001",
                "Descricao": "Oleo A",
                "Qtd": 2,
                "Unidade": "UN",
                "Peso": 500.0,
                "Destinatario": "Cliente A",
                "ROTA": "Sorocaba",
                "ChaveNFe": "35260600000000000000000000000000000000000041",
                "Status": "Autorizado o uso da NF-e",
            }
        ]
    )


def _build_summary() -> dict:
    return {
        "numero_carga": "CARGA-HIST-01",
        "motorista": "Joao Silva",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "02/07/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 500.0,
    }


def test_montar_painel_auditoria_com_impressoes() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            fechamento = get_fechamento_service()
            processed_df = _build_processed_df()
            from carregamentos.bootstrap import get_analise_operacional_service

            diagnostico = get_analise_operacional_service().analisar_lote_processado(processed_df)
            primeira = fechamento.executar_fechamento_veiculo(
                summary=_build_summary(),
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=diagnostico,
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.carregamento is not None

            fechamento.executar_reimpressao(
                primeira.carregamento.id,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
            )

            painel = get_historico_carregamento_service().montar_painel_auditoria(
                primeira.carregamento.id,
                excel_contexto="012 Romaneio Daniel.xlsx",
            )
            assert painel is not None
            assert painel.numero_carregamento == "000001"
            assert painel.estatisticas.total_nfs == 1
            assert painel.estatisticas.quantidade_reimpressoes == 1
            assert len(painel.nfs) == 1
            assert painel.nfs[0].cliente == "Cliente A"
            assert painel.nfs[0].excel_utilizado == "012 Romaneio Daniel.xlsx"
            assert len(painel.impressoes) >= 2
            assert painel.to_dict()["carregamento_id"] == primeira.carregamento.id
        finally:
            _teardown_sql_env()
