from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
from carregamentos.models.operacional import CenarioOperacional, ClassificacaoNfLote, DecisaoOperacional
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'analise.db').as_posix()}"
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
    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None


def _row(nf: str, chave: str, status: str = "Autorizado o uso da NF-e") -> dict:
    return {
        "NF": nf,
        "cProd": "P001",
        "Descricao": "Oleo A",
        "Qtd": 1,
        "Unidade": "UN",
        "Peso": 100.0,
        "Destinatario": "Cliente",
        "ROTA": "Rota 1",
        "ChaveNFe": chave,
        "Status": status,
    }


def _df(*rows: dict) -> pd.DataFrame:
    return pd.DataFrame(list(rows))


def _summary(numero: str = "CARGA-01") -> dict:
    return {
        "numero_carga": numero,
        "motorista": "Joao",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "02/07/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 100.0,
    }


def _persistir(df: pd.DataFrame, numero: str = "CARGA-01") -> None:
    fechamento = get_fechamento_service()
    analise = get_analise_operacional_service()
    diagnostico = analise.analisar_lote_processado(df)
    result = fechamento.executar_fechamento_veiculo(
        summary=_summary(numero),
        processed_df=df,
        current_user=None,
        gerar_minuta=True,
        gerar_romaneio=False,
        diagnostico=diagnostico,
        decisao=DecisaoOperacional.NOVO,
    )
    assert result.status == "primeira_impressao"


def test_lote_totalmente_novo() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            df = _df(_row("3001", "35260600000000000000000000000000000000000031"))
            diagnostico = get_analise_operacional_service().analisar_lote_processado(df)
            assert diagnostico.cenario == CenarioOperacional.NOVO
            assert diagnostico.nfs_novas == 1
            assert diagnostico.requer_decisao is False
            assert diagnostico.opcoes_decisao == [DecisaoOperacional.NOVO]
        finally:
            _teardown_sql_env()


def test_lote_totalmente_repetido_requer_reimpressao() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            df = _df(_row("3002", "35260600000000000000000000000000000000000032"))
            _persistir(df)
            diagnostico = get_analise_operacional_service().analisar_lote_processado(df)
            assert diagnostico.cenario == CenarioOperacional.REIMPRESSAO
            assert diagnostico.nfs_existentes == 1
            assert diagnostico.requer_decisao is True
            assert DecisaoOperacional.REIMPRIMIR in diagnostico.opcoes_decisao
        finally:
            _teardown_sql_env()


def test_lote_parcialmente_repetido_complementacao() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            existente = _df(_row("3003", "35260600000000000000000000000000000000000033"))
            _persistir(existente)
            misto = _df(
                _row("3003", "35260600000000000000000000000000000000000033"),
                _row("3004", "35260600000000000000000000000000000000000034"),
            )
            diagnostico = get_analise_operacional_service().analisar_lote_processado(misto)
            assert diagnostico.cenario == CenarioOperacional.COMPLEMENTACAO
            assert diagnostico.nfs_existentes == 1
            assert diagnostico.nfs_novas == 1
            assert DecisaoOperacional.COMPLEMENTAR in diagnostico.opcoes_decisao
        finally:
            _teardown_sql_env()


def test_nf_cancelada_bloqueia() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            df = _df(_row("3005", "35260600000000000000000000000000000000000035", "Cancelada"))
            diagnostico = get_analise_operacional_service().analisar_lote_processado(df)
            assert diagnostico.cenario == CenarioOperacional.NF_CANCELADA
            assert diagnostico.bloqueia_fechamento is True
        finally:
            _teardown_sql_env()


def test_conflito_multiplo_bloqueia() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            _persistir(_df(_row("3010", "35260600000000000000000000000000000000000040")), "CARGA-A")
            _persistir(_df(_row("3011", "35260600000000000000000000000000000000000041")), "CARGA-B")
            misto = _df(
                _row("3010", "35260600000000000000000000000000000000000040"),
                _row("3011", "35260600000000000000000000000000000000000041"),
            )
            diagnostico = get_analise_operacional_service().analisar_lote_processado(misto)
            assert diagnostico.cenario == CenarioOperacional.CONFLITO_MULTIPLO
            assert diagnostico.bloqueia_fechamento is True
        finally:
            _teardown_sql_env()


def test_filtrar_nfs_novas() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            service = get_analise_operacional_service()
            existente = _df(_row("3020", "35260600000000000000000000000000000000000050"))
            _persistir(existente)
            misto = _df(
                _row("3020", "35260600000000000000000000000000000000000050"),
                _row("3021", "35260600000000000000000000000000000000000051"),
            )
            diagnostico = service.analisar_lote_processado(misto)
            novos = service.filtrar_nfs_novas(misto, diagnostico)
            assert len(novos) == 1
            assert str(novos.iloc[0]["NF"]) == "3021"
        finally:
            _teardown_sql_env()
