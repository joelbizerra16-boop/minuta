from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pandas as pd

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
from carregamentos.models.auditoria_nf import NfAuditoriaCard, TipoOperacaoNf
from carregamentos.models.operacional import DecisaoOperacional
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.schema import ensure_full_schema


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'auditoria_nf.db').as_posix()}"
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


def _build_processed_df(nf: str = "1426799", chave: str = "35260600000000000000000000000000000000000099") -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "NF": nf,
                "cProd": "P001",
                "Descricao": "Oleo A",
                "Qtd": 1,
                "Unidade": "UN",
                "Peso": 412.0,
                "Destinatario": "DUGRAU MOTO PECAS LTDA ME",
                "ROTA": "Suzano",
                "ChaveNFe": chave,
                "Status": "Autorizado o uso da NF-e",
            }
        ]
    )


def _build_summary() -> dict:
    return {
        "numero_carga": "CARGA-AUD-01",
        "motorista": "Higor Carneiro",
        "placa": "GAZ3H44",
        "filial": "BRIDA",
        "data_saida": "02/07/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 412.0,
    }


def test_auditoria_nf_history_dataframe_colunas_e_ordem() -> None:
    from carregamentos.models.auditoria_nf import NfAuditoriaEvento, NfAuditoriaResumo, TipoOperacaoNf
    from carregamentos.ui.auditoria_nf_panel import build_history_dataframe

    card = NfAuditoriaCard(
        token="nf:1",
        nf="1426799",
        cliente="DUGRAU",
        situacao_atual="Ja utilizada anteriormente",
        quantidade_utilizacoes=2,
        ultima_utilizacao="02/07/2026 10:19",
        resumo=NfAuditoriaResumo(
            primeira_utilizacao="01/07/2026",
            ultima_utilizacao="02/07/2026",
            total_impressoes=1,
            total_reimpressoes=1,
            total_complementacoes=0,
            total_reentregas=0,
            carregamentos_distintos=1,
        ),
        eventos=[
            NfAuditoriaEvento(
                data="02/07/2026",
                hora="10:19",
                usuario="admin",
                operacao=TipoOperacaoNf.REIMPRESSAO,
                operacao_label="REIMPRESSAO",
                numero_carregamento="000027",
                motorista="Higor",
                rota="SUZANO",
                placa="GAZ3H44",
                status_carregamento="FINALIZADO",
                filial="BRIDA",
                tipo_operacao="ROTA",
                ordenacao=2.0,
            ),
            NfAuditoriaEvento(
                data="01/07/2026",
                hora="08:42",
                usuario="admin",
                operacao=TipoOperacaoNf.IMPRESSAO_ORIGINAL,
                operacao_label="IMPRESSAO ORIGINAL",
                numero_carregamento="000027",
                motorista="Higor",
                rota="SUZANO",
                placa="GAZ3H44",
                status_carregamento="FINALIZADO",
                filial="BRIDA",
                tipo_operacao="ROTA",
                ordenacao=1.0,
            ),
        ],
    )
    history_df = build_history_dataframe(card)
    assert list(history_df.columns) == [
        "Etapa",
        "Veiculo",
        "Placa",
        "Motorista",
        "Data",
        "Hora",
        "Usuario",
        "Carregamento",
        "IdCarga",
        "Rota",
        "Observacao",
    ]
    assert history_df.iloc[0]["Etapa"] == "Carregamento"
    assert history_df.iloc[1]["Etapa"] == "Reimpressao"
    assert history_df.iloc[0]["Motorista"] == "Higor"
    assert history_df.iloc[1]["Placa"] == "--"
    assert history_df.iloc[0]["Data"] == "01/07/2026"
    assert history_df.iloc[1]["Data"] == "02/07/2026"


def test_auditoria_nf_serialize_deserialize_processed_df() -> None:
    from carregamentos.ui.auditoria_nf_panel import (
        _deserialize_processed_df,
        _serialize_processed_df,
    )

    original = _build_processed_df()
    payload = _serialize_processed_df(original)

    assert isinstance(payload, tuple)
    assert len(payload) == 1
    first_row = payload[0]
    assert isinstance(first_row, tuple)
    assert all(isinstance(pair, tuple) and len(pair) == 2 for pair in first_row)
    assert any(pair[0] == "NF" and pair[1] == "1426799" for pair in first_row)

    restored = _deserialize_processed_df(payload)
    assert list(restored.columns) == list(original.columns)
    assert restored.iloc[0]["NF"] == original.iloc[0]["NF"]
    assert restored.iloc[0]["Destinatario"] == original.iloc[0]["Destinatario"]


def test_auditoria_nf_nova_sem_historico() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            processed_df = _build_processed_df("9999999", "35260600000000000000000000000000000000000100")
            auditoria = get_analise_operacional_service().montar_auditoria_nfs_lote(processed_df)
            assert len(auditoria.cards) == 1
            card = auditoria.cards[0]
            assert card.situacao_atual == "Nunca utilizada anteriormente"
            assert card.eventos == []
            assert card.quantidade_utilizacoes == 0
        finally:
            _teardown_sql_env()


def test_auditoria_nf_com_reimpressao_ordenada_desc() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            processed_df = _build_processed_df()
            analise = get_analise_operacional_service()
            fechamento = get_fechamento_service()
            diagnostico = analise.analisar_lote_processado(processed_df)
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
            analise.invalidar_cache()

            auditoria = analise.montar_auditoria_nfs_lote(processed_df)
            card = auditoria.cards[0]
            assert card.situacao_atual == "Ja utilizada anteriormente"
            assert len(card.eventos) >= 2
            assert card.eventos[0].ordenacao >= card.eventos[-1].ordenacao
            assert any(item.operacao == TipoOperacaoNf.IMPRESSAO_ORIGINAL for item in card.eventos)
            assert any(item.operacao == TipoOperacaoNf.REIMPRESSAO for item in card.eventos)
            assert card.resumo.total_reimpressoes >= 1
        finally:
            _teardown_sql_env()


def test_sql_auditoria_nf_parse_criado_em_tipos() -> None:
    from datetime import datetime, timezone

    from carregamentos.repository.sql_auditoria_nf_repository import SqlAuditoriaNfRepository

    aware = datetime(2026, 7, 2, 10, 19, tzinfo=timezone.utc)
    naive = datetime(2026, 7, 2, 10, 19)

    assert SqlAuditoriaNfRepository._parse_criado_em(None) is None
    assert SqlAuditoriaNfRepository._parse_criado_em(aware) == aware
    parsed_naive = SqlAuditoriaNfRepository._parse_criado_em(naive)
    assert parsed_naive is not None
    assert parsed_naive.tzinfo is not None
    parsed_string = SqlAuditoriaNfRepository._parse_criado_em("2026-07-02 10:19:00")
    assert parsed_string is not None
    assert parsed_string.hour == 10
    assert parsed_string.minute == 19
    assert parsed_string.tzinfo is not None


def test_extrato_nf_cache_roundtrip_sem_attribute_error() -> None:
    from datetime import datetime

    from carregamentos.repository.sql_auditoria_nf_repository import SqlAuditoriaNfRepository
    from carregamentos.ui.auditoria_nf_panel import _carregar_extrato_nf_cache, _restaurar_movimentacoes

    with tempfile.TemporaryDirectory() as tmp_dir:
        _setup_sql_env(Path(tmp_dir))
        try:
            processed_df = _build_processed_df()
            analise = get_analise_operacional_service()
            fechamento = get_fechamento_service()
            diagnostico = analise.analisar_lote_processado(processed_df)
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
            analise.invalidar_cache()
            _carregar_extrato_nf_cache.clear()

            movimentacoes = SqlAuditoriaNfRepository().buscar_extrato_movimentacoes_nf(
                numero_nf="1426799",
            )
            assert movimentacoes
            for item in movimentacoes:
                assert item.criado_em is None or isinstance(item.criado_em, datetime)

            payload = _carregar_extrato_nf_cache("roundtrip", "1426799", "")
            restauradas = _restaurar_movimentacoes(payload)
            assert restauradas
            for item in restauradas:
                assert item.criado_em is None or isinstance(item.criado_em, datetime)
        finally:
            _teardown_sql_env()
            _carregar_extrato_nf_cache.clear()
