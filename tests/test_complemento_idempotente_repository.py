from __future__ import annotations

import os
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
from sqlalchemy import event, select

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage, get_analise_operacional_service, get_fechamento_service
from carregamentos.models.carregamento import (
    MODALIDADE_VEICULO,
    STATUS_FINALIZADO,
    Carregamento,
    CarregamentoItem,
    utc_now_iso,
)
from carregamentos.models.operacional import DecisaoOperacional
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from core.settings import get_settings
from infrastructure.database import configure_database, get_engine, get_session_factory
from infrastructure.models.carregamento import ItemCarregamentoORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.schema import ensure_full_schema
from infrastructure.unit_of_work import UnitOfWork


def _setup_sql_env(data_dir: Path) -> None:
    os.environ["MINUTA_STORAGE_BACKEND"] = "sql"
    os.environ["MINUTA_DATA_ROOT"] = str(data_dir)
    os.environ["MINUTA_DATABASE_URL"] = f"sqlite:///{(data_dir / 'complemento.db').as_posix()}"
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


def _item(
    nf: str,
    cprod: str = "P001",
    *,
    chave: str | None = None,
    peso: float = 10.0,
    descricao: str = "Produto",
) -> CarregamentoItem:
    return CarregamentoItem(
        nf=nf,
        cprod=cprod,
        descricao=descricao,
        quantidade=1,
        unidade="UN",
        peso=peso,
        destinatario="Cliente",
        rota="R1",
        chave_nfe=chave or f"352606{nf.zfill(38)}"[:44],
        status_nf="OK",
    )


def _carregamento(usuario: UsuarioORM, itens: list[CarregamentoItem], numero: str, id_: int = 0) -> Carregamento:
    return Carregamento(
        id=id_,
        numero_carregamento=numero,
        data=date.today().isoformat(),
        hora="12:00:00",
        usuario=usuario.usuario,
        usuario_id=int(usuario.id),
        motorista="Motorista",
        placa="ABC1D23",
        filial="BRIDA",
        data_saida="--",
        quantidade_nf=len({i.nf for i in itens}),
        quantidade_itens=len(itens),
        peso_total=float(sum(i.peso for i in itens)),
        status=STATUS_FINALIZADO,
        modalidade=MODALIDADE_VEICULO,
        reentrega=False,
        minuta_pdf_path=None,
        romaneio_pdf_path=None,
        itens=itens,
        criado_em=utc_now_iso(),
    )


def _row(nf: str, chave: str, cprod: str = "P001", peso: float = 100.0) -> dict:
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


def _summary() -> dict:
    return {
        "numero_carga": "CARGA-COMP",
        "motorista": "Joao Silva",
        "placa": "ABC1D23",
        "filial": "BRIDA",
        "data_saida": "10/07/2026",
        "nf_count": 1,
        "item_count": 1,
        "peso_total": 100.0,
    }


class _SqlCounter:
    def __init__(self) -> None:
        self.inserts_item = 0
        self.updates_item = 0
        self.deletes_item = 0

    def listen(self, engine) -> None:
        event.listen(engine, "before_cursor_execute", self._on_cursor)

    def remove(self, engine) -> None:
        event.remove(engine, "before_cursor_execute", self._on_cursor)

    def _on_cursor(self, conn, cursor, statement, parameters, context, executemany):
        sql = " ".join(str(statement).lower().split())
        if "item_carregamento" not in sql:
            return
        if sql.startswith("insert into item_carregamento"):
            self.inserts_item += 1
        elif sql.startswith("update item_carregamento"):
            self.updates_item += 1
        elif sql.startswith("delete from item_carregamento"):
            self.deletes_item += 1


def test_novo_carregamento_somente_insert_itens() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        counter = _SqlCounter()
        engine = get_engine()
        counter.listen(engine)
        try:
            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                saved = repo._save_in_session(
                    uow.session,
                    _carregamento(usuario, [_item("1001"), _item("1002", "P002")], "000101"),
                )
                assert saved.id > 0
                assert len(saved.itens) == 2
            assert counter.inserts_item == 2
            assert counter.updates_item == 0
        finally:
            counter.remove(engine)
            _teardown_sql_env()


def test_complemento_idempotente_sem_insert_duplicado() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        counter = _SqlCounter()
        engine = get_engine()
        counter.listen(engine)
        try:
            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                saved = repo._save_in_session(
                    uow.session,
                    _carregamento(usuario, [_item("2001", descricao="Original")], "000201"),
                )
                cid = saved.id
                item_id = uow.session.scalars(
                    select(ItemCarregamentoORM.id).where(ItemCarregamentoORM.carregamento_id == cid)
                ).one()

            inserts_after_create = counter.inserts_item
            assert inserts_after_create == 1

            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                repo._save_in_session(
                    uow.session,
                    _carregamento(
                        usuario,
                        [_item("2001", descricao="Atualizado", peso=22.0)],
                        "000201",
                        id_=cid,
                    ),
                )
                row = uow.session.scalars(
                    select(ItemCarregamentoORM).where(ItemCarregamentoORM.carregamento_id == cid)
                ).one()
                assert int(row.id) == int(item_id)
                assert row.descricao == "Atualizado"
                assert float(row.peso) == 22.0

            assert counter.inserts_item == inserts_after_create
            assert counter.updates_item >= 1
        finally:
            counter.remove(engine)
            _teardown_sql_env()


def test_complemento_parcial_insere_somente_novos() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        counter = _SqlCounter()
        engine = get_engine()
        counter.listen(engine)
        try:
            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                saved = repo._save_in_session(
                    uow.session,
                    _carregamento(usuario, [_item("3001", "P001")], "000301"),
                )
                cid = saved.id
                existing_id = uow.session.scalars(
                    select(ItemCarregamentoORM.id).where(ItemCarregamentoORM.carregamento_id == cid)
                ).one()

            inserts_after_create = counter.inserts_item

            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                repo._save_in_session(
                    uow.session,
                    _carregamento(
                        usuario,
                        [_item("3001", "P001"), _item("3002", "P002")],
                        "000301",
                        id_=cid,
                    ),
                )
                rows = uow.session.scalars(
                    select(ItemCarregamentoORM)
                    .where(ItemCarregamentoORM.carregamento_id == cid)
                    .order_by(ItemCarregamentoORM.sequencia)
                ).all()
                assert len(rows) == 2
                assert int(rows[0].id) == int(existing_id)
                assert rows[0].numero_nf == "3001"
                assert rows[1].numero_nf == "3002"

            assert counter.inserts_item == inserts_after_create + 1
        finally:
            counter.remove(engine)
            _teardown_sql_env()


def test_complemento_multiplas_nfs_e_duas_execucoes_consecutivas() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        counter = _SqlCounter()
        engine = get_engine()
        counter.listen(engine)
        try:
            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                saved = repo._save_in_session(
                    uow.session,
                    _carregamento(
                        usuario,
                        [_item("4001", "A"), _item("4002", "B")],
                        "000401",
                    ),
                )
                cid = saved.id

            inserts_after_create = counter.inserts_item
            assert inserts_after_create == 2

            lista = [
                _item("4001", "A"),
                _item("4002", "B"),
                _item("4003", "C"),
                _item("4004", "D"),
            ]
            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                repo._save_in_session(
                    uow.session,
                    _carregamento(usuario, lista, "000401", id_=cid),
                )

            inserts_after_first_comp = counter.inserts_item
            assert inserts_after_first_comp == inserts_after_create + 2

            with UnitOfWork() as uow:
                usuario = uow.session.scalars(select(UsuarioORM).order_by(UsuarioORM.id)).first()
                repo = SqlCarregamentoRepository(uow.session)
                repo._save_in_session(
                    uow.session,
                    _carregamento(usuario, lista, "000401", id_=cid),
                )
                count = uow.session.scalar(
                    select(ItemCarregamentoORM.id).where(ItemCarregamentoORM.carregamento_id == cid)
                )
                rows = uow.session.scalars(
                    select(ItemCarregamentoORM).where(ItemCarregamentoORM.carregamento_id == cid)
                ).all()
                assert len(rows) == 4
                assert count is not None

            assert counter.inserts_item == inserts_after_first_comp
        finally:
            counter.remove(engine)
            _teardown_sql_env()


def test_lote_misto_fechamento_sem_integrity_error() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            base = pd.DataFrame([_row("5001", "35260600000000000000000000000000000000005001", "P001")])
            primeira = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=base,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=get_analise_operacional_service().analisar_lote_processado(base),
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.status == "primeira_impressao"

            misto = pd.DataFrame(
                [
                    _row("5001", "35260600000000000000000000000000000000005001", "P001"),
                    _row("5002", "35260600000000000000000000000000000000005002", "P002"),
                ]
            )
            result = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=misto,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=get_analise_operacional_service().analisar_lote_processado(misto),
                decisao=DecisaoOperacional.COMPLEMENTAR,
            )
            assert result.status == "complementacao"
            assert result.carregamento is not None
            assert len(result.carregamento.itens) == 2
            assert "IntegrityError" not in (result.message or "")
        finally:
            _teardown_sql_env()


def test_reimpressao_e_reentrega_sem_regressao() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        data_dir = Path(tmp_dir)
        _setup_sql_env(data_dir)
        try:
            fechamento = get_fechamento_service()
            df = pd.DataFrame([_row("6001", "35260600000000000000000000000000000000006001")])
            primeira = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=get_analise_operacional_service().analisar_lote_processado(df),
                decisao=DecisaoOperacional.NOVO,
            )
            assert primeira.carregamento is not None
            item_ids_antes = {
                (i.nf, i.cprod)
                for i in primeira.carregamento.itens
            }

            reimp = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=get_analise_operacional_service().analisar_lote_processado(df),
                decisao=DecisaoOperacional.REIMPRIMIR,
                confirmar_reimpressao=True,
            )
            assert reimp.status in {"reimpressao", "primeira_impressao"} or reimp.carregamento is not None

            reent = fechamento.executar_fechamento_veiculo(
                summary=_summary(),
                processed_df=df,
                current_user=None,
                gerar_minuta=True,
                gerar_romaneio=False,
                diagnostico=get_analise_operacional_service().analisar_lote_processado(df),
                decisao=DecisaoOperacional.REENTREGA,
                is_reentrega=True,
            )
            assert reent.status in {"reimpressao", "reentrega"} or reent.carregamento is not None
            assert reent.carregamento is not None
            assert {(i.nf, i.cprod) for i in reent.carregamento.itens} == item_ids_antes
        finally:
            _teardown_sql_env()
