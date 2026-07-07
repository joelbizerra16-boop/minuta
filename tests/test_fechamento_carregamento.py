from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
from carregamentos.models.carregamento import MODALIDADE_VEICULO, STATUS_FINALIZADO
from carregamentos.models.operacional import DecisaoOperacional
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'fechamento.db').as_posix()}"
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
    carregamentos_bootstrap._repository = None
    carregamentos_bootstrap._fechamento_service = None
    carregamentos_bootstrap._analise_operacional_service = None
    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None


def _build_processed_df(nf: str = "2001", chave: str = "35260600000000000000000000000000000000000011") -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "NF": nf,
                "cProd": "P001",
                "Descricao": "Oleo A",
                "Qtd": 2,
                "Unidade": "UN",
                "Peso": 500.0,
                "Destinatario": "Cliente A",
                "ROTA": "Rota 1",
                "ChaveNFe": chave,
                "Status": "Autorizado o uso da NF-e",
            }
        ]
    )


def _build_summary(numero: str = "CARGA-FECH-01") -> dict:
    return {
        "numero_carga": numero,
        "motorista": "Joao Silva",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "29/06/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 500.0,
    }


def _analisar(df: pd.DataFrame):
    return get_analise_operacional_service().analisar_lote_processado(df)


def test_primeira_impressao_persiste_antes_do_pdf() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            summary = _build_summary()
            processed_df = _build_processed_df()
            diagnostico = _analisar(processed_df)
            result = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=diagnostico,
                decisao=DecisaoOperacional.NOVO,
            )
            assert result.status == "primeira_impressao"
            assert result.carregamento is not None
            assert result.carregamento.id > 0
            assert result.carregamento.numero_carregamento == "000001"
            assert len(result.carregamento.itens) == 1
            assert result.carregamento.status == STATUS_FINALIZADO
            assert result.carregamento.modalidade == MODALIDADE_VEICULO

            reloaded = fechamento._repository.get_by_numero("000001")
            assert reloaded is not None
            assert len(reloaded.itens) == 1
            assert reloaded.quantidade_impressoes == 1
        finally:
            _teardown_sql_env()


def test_mesmas_notas_sem_decisao_bloqueia_fechamento() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            summary = _build_summary()
            processed_df = _build_processed_df()
            primeira = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(processed_df),
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.carregamento is not None

            segunda = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(processed_df),
            )
            assert segunda.status == "invalid"
            assert "Decisao operacional pendente" in (segunda.message or "")
        finally:
            _teardown_sql_env()


def test_reimpressao_preserva_carregamento_existente() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            summary = _build_summary()
            processed_df = _build_processed_df()
            primeira = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(processed_df),
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.carregamento is not None

            reimpressao = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(processed_df),
                decisao=DecisaoOperacional.REIMPRIMIR,
            )
            assert reimpressao.status == "reimpressao"
            assert reimpressao.carregamento is not None
            assert reimpressao.carregamento.id == primeira.carregamento.id
            assert reimpressao.carregamento.quantidade_impressoes == 2
        finally:
            _teardown_sql_env()


def test_complementacao_adiciona_somente_nfs_novas() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            base_df = _build_processed_df("2001", "35260600000000000000000000000000000000000011")
            fechamento.executar_fechamento_veiculo(
                summary=_build_summary(),
                processed_df=base_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(base_df),
                decisao=DecisaoOperacional.NOVO,
            )

            misto_df = pd.DataFrame(
                [
                    {
                        "NF": "2001",
                        "cProd": "P001",
                        "Descricao": "Oleo A",
                        "Qtd": 2,
                        "Unidade": "UN",
                        "Peso": 500.0,
                        "Destinatario": "Cliente A",
                        "ROTA": "Rota 1",
                        "ChaveNFe": "35260600000000000000000000000000000000000011",
                        "Status": "Autorizado o uso da NF-e",
                    },
                    {
                        "NF": "2002",
                        "cProd": "P002",
                        "Descricao": "Oleo B",
                        "Qtd": 1,
                        "Unidade": "UN",
                        "Peso": 300.0,
                        "Destinatario": "Cliente B",
                        "ROTA": "Rota 2",
                        "ChaveNFe": "35260600000000000000000000000000000000000012",
                        "Status": "Autorizado o uso da NF-e",
                    },
                ]
            )
            diagnostico = _analisar(misto_df)
            result = fechamento.executar_fechamento_veiculo(
                summary=_build_summary("CARGA-FECH-02"),
                processed_df=misto_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=diagnostico,
                decisao=DecisaoOperacional.COMPLEMENTAR,
            )
            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert result.carregamento.numero_carregamento == "000001"
            assert len(result.carregamento.itens) == 2
        finally:
            _teardown_sql_env()


def test_reimpressao_explicita_preserva_carregamento_anterior() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            summary = _build_summary()
            processed_df = _build_processed_df()
            primeira = fechamento.executar_fechamento_veiculo(
                summary=summary,
                processed_df=processed_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(processed_df),
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.carregamento is not None

            reimpressao = fechamento.executar_reimpressao(
                primeira.carregamento.id,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
            )
            assert reimpressao.status == "reimpressao"
            assert reimpressao.carregamento is not None
            assert reimpressao.carregamento.id == primeira.carregamento.id
            assert reimpressao.carregamento.quantidade_impressoes == 2
        finally:
            _teardown_sql_env()
