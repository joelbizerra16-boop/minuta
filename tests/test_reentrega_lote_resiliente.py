from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import pandas as pd
import pytest
from sqlalchemy import select

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
from carregamentos.models.operacional import ClassificacaoOperacionalNf, DecisaoOperacional
from carregamentos.services.nf_operacional_classifier import NfOperacionalClassifier
from carregamentos.services.validacao_item_carregamento import filtrar_itens_novos_para_insercao
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.models.constants import AUDIT_EVENTO_DECISAO_OPERACIONAL
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'reentrega.db').as_posix()}"
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


def _row(
    nf: str,
    chave: str,
    cprod: str = "P001",
    peso: float = 500.0,
) -> dict:
    return {
        "NF": nf,
        "cProd": cprod,
        "Descricao": "Oleo A",
        "Qtd": 2,
        "Unidade": "UN",
        "Peso": peso,
        "Destinatario": "Cliente A",
        "ROTA": "Rota 1",
        "ChaveNFe": chave,
        "Status": "Autorizado o uso da NF-e",
    }


def _summary(numero: str = "CARGA-REENTREGA") -> dict:
    return {
        "numero_carga": numero,
        "motorista": "Joao Silva",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "08/07/2026",
        "nf_count": 2,
        "item_count": 2,
        "peso_total": 800.0,
    }


def _analisar(df: pd.DataFrame):
    return get_analise_operacional_service().analisar_lote_processado(df)


def test_lote_misto_com_reentrega_nao_gera_integrity_error() -> None:
    """Reproduz o incidente de producao: lote misto com NF ja existente no carregamento."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            nf_existente = "1428910"
            chave_existente = "35260700846804000106550010014289101517466131"
            base_df = pd.DataFrame([_row(nf_existente, chave_existente, "123737", 102.6491)])
            fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=base_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(base_df),
                decisao=DecisaoOperacional.NOVO,
            )

            misto_df = pd.DataFrame(
                [
                    _row(nf_existente, chave_existente, "123737", 102.6491),
                    _row("1428911", "35260700846804000106550010014289111517466132", "123738", 200.0),
                ]
            )
            diagnostico = _analisar(misto_df)
            result = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=misto_df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=diagnostico,
                decisao=DecisaoOperacional.COMPLEMENTAR,
            )
            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert len(result.carregamento.itens) == 2
            assert "IntegrityError" not in (result.message or "")
            assert result.relatorio_lote is not None
            assert result.relatorio_lote.reentregas >= 1 or result.relatorio_lote.duplicidades >= 1
        finally:
            _teardown_sql_env()


def test_complementacao_somente_reentregas_sem_insert() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            df = pd.DataFrame([_row("5001", "35260600000000000000000000000000000000000091")])
            fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(df),
                decisao=DecisaoOperacional.NOVO,
            )
            result = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(df),
                decisao=DecisaoOperacional.COMPLEMENTAR,
            )
            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert len(result.carregamento.itens) == 1
        finally:
            _teardown_sql_env()


def test_classificador_marca_duplicidade_no_mesmo_carregamento() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            repo = fechamento._repository
            df = pd.DataFrame([_row("6001", "35260600000000000000000000000000000000000101")])
            fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(df),
                decisao=DecisaoOperacional.NOVO,
            )
            carregamento = repo.get_by_numero("000001")
            assert carregamento is not None
            diagnostico = _analisar(df)
            classifier = NfOperacionalClassifier()
            diagnosticos = classifier.classificar_lote(
                df,
                diagnostico,
                DecisaoOperacional.COMPLEMENTAR,
                carregamento,
            )
            assert len(diagnosticos) == 1
            assert diagnosticos[0].classificacao == ClassificacaoOperacionalNf.DUPLICIDADE
        finally:
            _teardown_sql_env()


def test_validacao_preventiva_detecta_duplicidade_logica() -> None:
    from carregamentos.models.carregamento import Carregamento, CarregamentoItem, MODALIDADE_VEICULO, STATUS_FINALIZADO, utc_now_iso
    from carregamentos.services.validacao_item_carregamento import filtrar_itens_novos_para_insercao

    item = CarregamentoItem(
        nf="100",
        cprod="X",
        descricao="",
        quantidade=1,
        unidade="UN",
        peso=1,
        destinatario="",
        rota="",
    )
    carregamento = Carregamento(
        id=1,
        numero_carregamento="000001",
        data="2026-07-08",
        hora="10:00",
        usuario="test",
        usuario_id=1,
        motorista="--",
        placa="--",
        filial="BRIDA",
        data_saida="--",
        quantidade_nf=1,
        quantidade_itens=1,
        peso_total=1,
        status=STATUS_FINALIZADO,
        modalidade=MODALIDADE_VEICULO,
        reentrega=False,
        minuta_pdf_path=None,
        romaneio_pdf_path=None,
        itens=[item],
        criado_em=utc_now_iso(),
    )
    novos, duplicados = filtrar_itens_novos_para_insercao(carregamento, [item])
    assert novos == []
    assert len(duplicados) == 1


def test_auditoria_decisao_operacional_registrada_na_complementacao() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            base = pd.DataFrame([_row("7001", "35260600000000000000000000000000000000000111")])
            fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=base,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(base),
                decisao=DecisaoOperacional.NOVO,
            )
            misto = pd.DataFrame(
                [
                    _row("7001", "35260600000000000000000000000000000000000111"),
                    _row("7002", "35260600000000000000000000000000000000000112"),
                ]
            )
            fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=misto,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=_analisar(misto),
                decisao=DecisaoOperacional.COMPLEMENTAR,
            )
            with get_engine().connect() as conn:
                count = conn.execute(
                    select(EventoAuditoriaORM).where(
                        EventoAuditoriaORM.evento == AUDIT_EVENTO_DECISAO_OPERACIONAL
                    )
                ).fetchall()
            assert len(count) >= 1
            meta = json.loads(count[0].metadados_json or "{}")
            assert "decisao" in meta
            assert "situacao_anterior" in meta
            assert "situacao_posterior" in meta
        finally:
            _teardown_sql_env()
