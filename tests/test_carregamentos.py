from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_carregamento_service
from carregamentos.models.carregamento import (
    MODALIDADE_BALCAO,
    MODALIDADE_VEICULO,
    STATUS_FINALIZADO,
    STATUS_FINALIZADO_R,
    CarregamentoFiltro,
)
from carregamentos.services.nf_validation import localizar_nf_no_lote
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
    )
    ensure_full_schema()
    configure_auth_storage(data_dir)
    configure_carregamentos_storage(data_dir)


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
                "ROTA": "Rota 1",
                "ChaveNFe": "35260600000000000000000000000000000000000001",
                "Status": "Autorizado o uso da NF-e",
            },
            {
                "NF": "1002",
                "cProd": "P002",
                "Descricao": "Oleo B",
                "Qtd": 1,
                "Unidade": "UN",
                "Peso": 1000.0,
                "Destinatario": "Cliente B",
                "ROTA": "Rota 2",
                "ChaveNFe": "35260600000000000000000000000000000000000002",
                "Status": "Autorizado o uso da NF-e",
            },
        ]
    )


def _build_summary() -> dict:
    return {
        "numero_carga": "CARGA-300",
        "motorista": "Joao Silva",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "29/06/2026",
        "nf_count": 2,
        "item_count": 2,
        "peso_total": 1500.0,
    }


def _teardown_sql_env() -> None:
    import carregamentos.bootstrap as carregamentos_bootstrap
    import infrastructure.database as db_module

    carregamentos_bootstrap.get_carregamento_service.cache_clear()
    carregamentos_bootstrap._repository = None
    get_engine().dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None


def test_nf_inedita_finalizacao_veiculo() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()

        saved = service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=b"%PDF-romaneio%",
        )
        assert saved is not None
        assert saved.status == STATUS_FINALIZADO
        assert saved.modalidade == MODALIDADE_VEICULO
        assert saved.reentrega is False
        print("nf inedita OK")

    _run_test(_case)


def _run_test(fn) -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fn(data_dir)
        finally:
            _teardown_sql_env()


def test_nf_duplicada_bloqueia_sem_reentrega() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()
        service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
        )

        conflitos = service.validar_conflitos_nf(processed_df)
        assert len(conflitos) >= 1

        try:
            service.register_from_processing(
                summary={**summary, "numero_carga": "CARGA-301"},
                processed_df=processed_df,
                current_user=None,
                carregamento_pdf=b"%PDF-minuta%",
                romaneio_pdf=None,
            )
            raise AssertionError("deveria bloquear duplicidade")
        except ValueError:
            pass
        print("nf duplicada bloqueio OK")

    _run_test(_case)


def test_reentrega_autorizada() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()
        service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
        )

        reentrega = service.register_from_processing(
            summary={**summary, "numero_carga": "CARGA-302"},
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
            is_reentrega=True,
        )
        assert reentrega is not None
        assert reentrega.status == STATUS_FINALIZADO_R
        assert reentrega.reentrega is True
        print("reentrega autorizada OK")

    _run_test(_case)


def test_entrega_balcao() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()

        balcao = service.register_entrega_balcao(
            termo_busca="1001",
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            standalone_balcao=True,
        )
        assert balcao is not None
        assert balcao.modalidade == MODALIDADE_BALCAO
        assert balcao.motorista == "--"
        assert balcao.placa == "--"
        assert balcao.numero_carregamento == "BALCAO-1001"
        assert len(balcao.itens) == 1
        print("entrega balcao OK")

    _run_test(_case)


def test_pesquisa_inteligente() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()
        saved = service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
        )
        assert saved is not None

        by_nf = service.search_itens_listagem(
            CarregamentoFiltro(data_inicial=saved.data, data_final=saved.data, termo_pesquisa="1001")
        )
        by_chave = service.search_itens_listagem(
            CarregamentoFiltro(
                data_inicial=saved.data,
                data_final=saved.data,
                termo_pesquisa="35260600000000000000000000000000000000000001",
            )
        )
        by_motorista = service.search_itens_listagem(
            CarregamentoFiltro(data_inicial=saved.data, data_final=saved.data, termo_pesquisa="joao")
        )
        assert len(by_nf.linhas) == 1
        assert by_nf.carregamentos_distintos == 1
        assert len(by_chave.linhas) == 1
        assert by_chave.carregamentos_distintos == 1
        assert len(by_motorista.linhas) == 2
        assert by_motorista.carregamentos_distintos == 1
        print("pesquisa inteligente OK")

    _run_test(_case)


def test_carregamentos_distintos_por_nf() -> None:
    def _case(_: Path) -> None:
        service = get_carregamento_service()
        processed_df = _build_processed_df()
        summary = _build_summary()
        primeiro = service.register_from_processing(
            summary=summary,
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
        )
        assert primeiro is not None
        service.register_from_processing(
            summary={**summary, "numero_carga": "CARGA-301"},
            processed_df=processed_df,
            current_user=None,
            carregamento_pdf=b"%PDF-minuta%",
            romaneio_pdf=None,
            is_reentrega=True,
        )

        listagem = service.search_itens_listagem(
            CarregamentoFiltro(
                data_inicial=primeiro.data,
                data_final=primeiro.data,
                termo_pesquisa="1001",
            )
        )
        assert len(listagem.linhas) == 2
        assert listagem.carregamentos_distintos == 2
        print("carregamentos distintos por nf OK")

    _run_test(_case)


def test_localizar_nf_no_lote() -> None:
    processed_df = _build_processed_df()
    encontrado = localizar_nf_no_lote(processed_df, "1002")
    assert len(encontrado) == 1
    assert encontrado.iloc[0]["NF"] == "1002"
    print("localizar nf OK")


if __name__ == "__main__":
    test_nf_inedita_finalizacao_veiculo()
    test_nf_duplicada_bloqueia_sem_reentrega()
    test_reentrega_autorizada()
    test_entrega_balcao()
    test_pesquisa_inteligente()
    test_carregamentos_distintos_por_nf()
    test_localizar_nf_no_lote()
    print("All carregamento control tests passed")
