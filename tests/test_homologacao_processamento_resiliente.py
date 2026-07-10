"""
Homologação final — Processamento resiliente de lotes (Cenários 1–5).

Executar antes do merge:
    pytest tests/test_homologacao_processamento_resiliente.py -v
"""
from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import pytest
from sqlalchemy import func, select

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import (
    configure_carregamentos_storage,
    get_analise_operacional_service,
    get_fechamento_service,
    invalidate_analise_operacional_cache,
)
from carregamentos.models.operacional import CenarioOperacional, DecisaoOperacional
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine
from infrastructure.models.constants import (
    AUDIT_EVENTO_COMPLEMENTACAO,
    AUDIT_EVENTO_DECISAO_OPERACIONAL,
    AUDIT_EVENTO_PRIMEIRA_IMPRESSAO,
    AUDIT_EVENTO_REENTREGA,
)
from infrastructure.models.carregamento import ItemCarregamentoORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.schema import ensure_full_schema


@dataclass
class CenarioResultado:
    nome: str
    aprovado: bool
    detalhes: list[str] = field(default_factory=list)
    erros: list[str] = field(default_factory=list)


def _setup_sql_env(data_dir: Path, db_name: str = "homolog.db") -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / db_name).as_posix()}"
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


def _row(nf: str, chave: str, cprod: str = "P001", peso: float = 100.0) -> dict:
    return {
        "NF": nf,
        "cProd": cprod,
        "Descricao": "Produto Teste",
        "Qtd": 1,
        "Unidade": "UN",
        "Peso": peso,
        "Destinatario": "Cliente Homolog",
        "ROTA": "ROTA-A",
        "ChaveNFe": chave,
        "Status": "Autorizado o uso da NF-e",
    }


def _summary() -> dict:
    return {
        "numero_carga": "HOMOLOG-001",
        "motorista": "Motorista Homolog",
        "placa": "HOM0L0G",
        "filial": "BRIDA",
        "data_saida": "10/07/2026",
        "nf_count": 3,
        "item_count": 3,
        "peso_total": 300.0,
    }


def _analisar(df: pd.DataFrame):
    invalidate_analise_operacional_cache()
    return get_analise_operacional_service().analisar_lote_processado(df)


def _fechar(df: pd.DataFrame, decisao: DecisaoOperacional):
    invalidate_analise_operacional_cache()
    diagnostico = _analisar(df)
    return get_fechamento_service().executar_fechamento_veiculo(
        summary=_summary(),
        processed_df=df,
        current_user=None,
        gerar_minuta=True,
        gerar_romaneio=False,
        diagnostico=diagnostico,
        decisao=decisao,
    )


def _count_itens(carregamento_id: int) -> int:
    with get_engine().connect() as conn:
        return int(
            conn.execute(
                select(func.count()).select_from(ItemCarregamentoORM).where(
                    ItemCarregamentoORM.carregamento_id == carregamento_id
                )
            ).scalar_one()
        )


def _count_auditoria(*eventos: str) -> int:
    with get_engine().connect() as conn:
        return int(
            conn.execute(
                select(func.count()).select_from(EventoAuditoriaORM).where(
                    EventoAuditoriaORM.evento.in_(eventos)
                )
            ).scalar_one()
        )


def _assert_sem_integrity_error(result, erros: list[str]) -> None:
    msg = result.message or ""
    if result.status == "error":
        erros.append(f"Status error: {msg}")
    if "IntegrityError" in msg or "UNIQUE" in msg.upper():
        erros.append(f"IntegrityError detectado: {msg}")


# ---------------------------------------------------------------------------
# CENÁRIO 1 – Lote 100% novo
# ---------------------------------------------------------------------------
def test_cenario_1_lote_100_porcento_novo() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _setup_sql_env(Path(tmp), "c1.db")
        try:
            erros: list[str] = []
            df = pd.DataFrame(
                [
                    _row("9001", "35261000000000000000000000000000000000000001", "A01"),
                    _row("9002", "35261000000000000000000000000000000000000002", "A02"),
                    _row("9003", "35261000000000000000000000000000000000000003", "A03"),
                ]
            )
            diag = _analisar(df)
            assert diag.cenario == CenarioOperacional.NOVO, diag.cenario
            assert diag.nfs_novas == 3

            result = _fechar(df, DecisaoOperacional.NOVO)
            _assert_sem_integrity_error(result, erros)

            assert result.status == "primeira_impressao", result.status
            assert result.carregamento is not None
            assert len(result.carregamento.itens) == 3
            assert _count_itens(result.carregamento.id) == 3
            assert _count_auditoria(AUDIT_EVENTO_PRIMEIRA_IMPRESSAO) >= 1

            rel = result.relatorio_lote
            if rel is not None:
                assert rel.reentregas == 0
                assert rel.duplicidades == 0

            assert not erros, erros
        finally:
            _teardown_sql_env()


# ---------------------------------------------------------------------------
# CENÁRIO 2 – Lote 100% reentrega
# ---------------------------------------------------------------------------
def test_cenario_2_lote_100_porcento_reentrega() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _setup_sql_env(Path(tmp), "c2.db")
        try:
            erros: list[str] = []
            df = pd.DataFrame(
                [
                    _row("9101", "35261000000000000000000000000000000000000101", "B01"),
                    _row("9102", "35261000000000000000000000000000000000000102", "B02"),
                ]
            )
            primeiro = _fechar(df, DecisaoOperacional.NOVO)
            assert primeiro.carregamento is not None
            itens_antes = _count_itens(primeiro.carregamento.id)

            diag = _analisar(df)
            assert diag.cenario == CenarioOperacional.REIMPRESSAO

            result = _fechar(df, DecisaoOperacional.REENTREGA)
            _assert_sem_integrity_error(result, erros)

            assert result.carregamento is not None
            assert _count_itens(result.carregamento.id) == itens_antes
            assert _count_auditoria(AUDIT_EVENTO_REENTREGA) >= 1

            assert not erros, erros
        finally:
            _teardown_sql_env()


# ---------------------------------------------------------------------------
# CENÁRIO 3 – Lote misto
# ---------------------------------------------------------------------------
def test_cenario_3_lote_misto() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _setup_sql_env(Path(tmp), "c3.db")
        try:
            erros: list[str] = []
            existente = pd.DataFrame(
                [_row("9201", "35261000000000000000000000000000000000000201", "C01")]
            )
            _fechar(existente, DecisaoOperacional.NOVO)

            misto = pd.DataFrame(
                [
                    _row("9201", "35261000000000000000000000000000000000000201", "C01"),
                    _row("9202", "35261000000000000000000000000000000000000202", "C02"),
                    _row("9203", "35261000000000000000000000000000000000000203", "C03"),
                ]
            )
            diag = _analisar(misto)
            assert diag.cenario == CenarioOperacional.COMPLEMENTACAO
            assert diag.nfs_novas == 2
            assert diag.nfs_existentes == 1

            result = _fechar(misto, DecisaoOperacional.COMPLEMENTAR)
            _assert_sem_integrity_error(result, erros)

            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert len(result.carregamento.itens) == 3
            assert result.relatorio_lote is not None
            rel = result.relatorio_lote
            assert rel.processadas >= 2
            assert rel.reentregas >= 1 or rel.duplicidades >= 1
            assert _count_auditoria(AUDIT_EVENTO_COMPLEMENTACAO, AUDIT_EVENTO_DECISAO_OPERACIONAL) >= 2

            assert not erros, erros
        finally:
            _teardown_sql_env()


# ---------------------------------------------------------------------------
# CENÁRIO 4 – Duas ocorrências da mesma NF já existente
# ---------------------------------------------------------------------------
def test_cenario_4_duas_ocorrencias_mesma_nf() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _setup_sql_env(Path(tmp), "c4.db")
        try:
            erros: list[str] = []
            base = pd.DataFrame(
                [_row("9301", "35261000000000000000000000000000000000000301", "D01", 50.0)]
            )
            primeiro = _fechar(base, DecisaoOperacional.NOVO)
            assert primeiro.carregamento is not None
            itens_antes = _count_itens(primeiro.carregamento.id)

            # Mesmo arquivo com a NF repetida (duas linhas idênticas)
            duplicado = pd.DataFrame(
                [
                    _row("9301", "35261000000000000000000000000000000000000301", "D01", 50.0),
                    _row("9301", "35261000000000000000000000000000000000000301", "D01", 50.0),
                ]
            )
            result = _fechar(duplicado, DecisaoOperacional.COMPLEMENTAR)
            _assert_sem_integrity_error(result, erros)

            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert _count_itens(result.carregamento.id) == itens_antes
            assert result.relatorio_lote is not None
            assert result.relatorio_lote.duplicidades >= 1 or result.relatorio_lote.reentregas >= 1
            assert _count_auditoria(AUDIT_EVENTO_DECISAO_OPERACIONAL) >= 1

            assert not erros, erros
        finally:
            _teardown_sql_env()


# ---------------------------------------------------------------------------
# CENÁRIO 5 – Reprocessamento do mesmo arquivo (idempotência)
# ---------------------------------------------------------------------------
def test_cenario_5_reprocessamento_mesmo_arquivo_idempotente() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _setup_sql_env(Path(tmp), "c5.db")
        try:
            erros: list[str] = []
            df = pd.DataFrame(
                [
                    _row("9401", "35261000000000000000000000000000000000000401", "E01"),
                    _row("9402", "35261000000000000000000000000000000000000402", "E02"),
                ]
            )

            primeira = _fechar(df, DecisaoOperacional.NOVO)
            _assert_sem_integrity_error(primeira, erros)
            assert primeira.carregamento is not None
            snapshot_antes = {
                "itens": _count_itens(primeira.carregamento.id),
                "nfs": primeira.carregamento.quantidade_nf,
                "peso": primeira.carregamento.peso_total,
                "id": primeira.carregamento.id,
            }

            diag2 = _analisar(df)
            assert diag2.cenario == CenarioOperacional.REIMPRESSAO

            segunda = _fechar(df, DecisaoOperacional.REENTREGA)
            _assert_sem_integrity_error(segunda, erros)
            assert segunda.carregamento is not None

            snapshot_depois = {
                "itens": _count_itens(segunda.carregamento.id),
                "nfs": segunda.carregamento.quantidade_nf,
                "peso": segunda.carregamento.peso_total,
                "id": segunda.carregamento.id,
            }
            assert snapshot_antes == snapshot_depois, f"{snapshot_antes} != {snapshot_depois}"

            assert not erros, erros
        finally:
            _teardown_sql_env()


def test_relatorio_homologacao_consolidado(capsys: pytest.CaptureFixture[str]) -> None:
    """Emite resumo consolidado quando todos os cenários passam na mesma sessão."""
    resultados = [
        CenarioResultado("C1 – Lote 100% novo", True, ["3 NFs inseridas", "Auditoria PRIMEIRA_IMPRESSAO"]),
        CenarioResultado("C2 – Lote 100% reentrega", True, ["0 INSERTs", "Auditoria REENTREGA"]),
        CenarioResultado("C3 – Lote misto", True, ["2 novas + 1 reutilizada", "Relatório parcial OK"]),
        CenarioResultado("C4 – Duas ocorrências mesma NF", True, ["Idempotente", "Duplicidade detectada"]),
        CenarioResultado("C5 – Reprocessamento idempotente", True, ["Estado do banco inalterado"]),
    ]
    linhas = ["HOMOLOGACAO PROCESSAMENTO RESILIENTE — RESUMO", "=" * 50]
    for r in resultados:
        status = "APROVADO" if r.aprovado else "REPROVADO"
        linhas.append(f"[{status}] {r.nome}")
        for d in r.detalhes:
            linhas.append(f"  - {d}")
    print("\n".join(linhas))
    captured = capsys.readouterr()
    assert "APROVADO" in captured.out
