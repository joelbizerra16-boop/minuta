from __future__ import annotations

import tempfile
from datetime import date, time, timedelta
from pathlib import Path
from unittest.mock import patch

from sqlalchemy import func, select

from auth.bootstrap import configure_auth_storage
from carregamentos.bootstrap import configure_carregamentos_storage
from carregamentos.services.retencao_automatica_service import RetencaoAutomaticaService
from core.retention_policy import retention_days_before_today
from core.startup_retention import reset_startup_retention_flag, run_startup_retention_once
from infrastructure.database import configure_database, get_engine
from infrastructure.models.carregamento import CarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.schema import ensure_full_schema
from infrastructure.unit_of_work import UnitOfWork
from tests.test_execucao_retencao import _build_services
from tests.test_gestao_retencao import _seed_retencao_dataset


def _reset_db_module() -> None:
    import infrastructure.database as db_module

    if db_module._engine is not None:
        db_module._engine.dispose()
    db_module._engine = None
    db_module._session_factory = None
    db_module._data_root = None
    db_module._pdf_storage_dir = None
    db_module._xml_storage_dir = None


def _configure_bootstrap(tmp_path: Path, db_name: str) -> Path:
    db_path = tmp_path / db_name
    pdf_dir = tmp_path / "documentos"
    xml_dir = tmp_path / "xml"
    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=tmp_path,
        pdf_storage_dir=pdf_dir,
        xml_storage_dir=xml_dir,
    )
    ensure_full_schema()
    configure_auth_storage(tmp_path)
    configure_carregamentos_storage(tmp_path)
    return pdf_dir


def test_retencao_automatica_sem_pacotes_elegiveis() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        _configure_bootstrap(tmp_path, "auto_sem_elegiveis.db")

        hoje = date.today()
        with UnitOfWork() as uow:
            from auth.repository.perfil_seed import seed_perfis

            seed_perfis(uow.session)
            usuario = UsuarioORM(
                uuid="00000000-0000-4000-8000-000000000099",
                nome="Teste",
                usuario="teste",
                senha_hash="pbkdf2_sha256$100000$aaaaaaaaaaaaaaaa$bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
                perfil_id=2,
                perfil="OPERADOR",
            )
            uow.session.add(usuario)
            uow.session.flush()
            uow.session.add(
                CarregamentoORM(
                    numero_carregamento="000001",
                    usuario_id=int(usuario.id),
                    data=hoje,
                    hora=time(9, 0),
                    motorista="Motorista",
                    placa="ABC1D23",
                    modalidade="VEICULO",
                    status="FINALIZADO",
                )
            )

        service = RetencaoAutomaticaService()
        resultado = service.executar()

        assert resultado.executado is False
        assert "Nenhum pacote elegivel encontrado." in resultado.mensagem
        assert resultado.pacotes_removidos == 0
        assert resultado.painel_atualizado is not None
        assert resultado.painel_atualizado.pacote.carregamentos == 0

        get_engine().dispose()
        _reset_db_module()


def test_retencao_automatica_remove_pacotes_aptos() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        pdf_dir = tmp_path / "documentos"
        db_path = tmp_path / "auto_remove.db"
        _seed_retencao_dataset(db_path, pdf_dir)
        _configure_bootstrap(tmp_path, "auto_remove.db")

        gestao, simulacao, execucao = _build_services(pdf_dir)
        service = RetencaoAutomaticaService(
            gestao_dados_service=gestao,
            simulacao_service=simulacao,
            execucao_service=execucao,
        )
        resultado = service.executar()

        assert resultado.executado is True
        assert resultado.pacotes_removidos == 1
        assert resultado.pacotes_mantidos == 0
        assert resultado.espaco_recuperado_bytes >= 0
        assert resultado.painel_atualizado is not None
        assert resultado.painel_atualizado.pacote.carregamentos == 0

        with UnitOfWork() as uow:
            total = int(uow.session.scalar(select(func.count()).select_from(CarregamentoORM)) or 0)
        assert total == 1

        get_engine().dispose()
        _reset_db_module()


def test_retencao_automatica_mantem_pacote_critico() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        pdf_dir = tmp_path / "documentos"
        db_path = tmp_path / "auto_critico.db"
        _seed_retencao_dataset(db_path, pdf_dir)
        _configure_bootstrap(tmp_path, "auto_critico.db")

        data_critica = date.today() - timedelta(days=retention_days_before_today() + 5)
        with UnitOfWork() as uow:
            usuario_id = int(
                uow.session.scalars(select(UsuarioORM.id).order_by(UsuarioORM.id)).first() or 0
            )
            critico = CarregamentoORM(
                numero_carregamento="000099",
                usuario_id=usuario_id,
                data=data_critica,
                hora=time(8, 0),
                motorista="Motorista",
                placa="ZZZ9Z99",
                modalidade="VEICULO",
                status="FINALIZADO",
            )
            uow.session.add(critico)
            uow.session.flush()
            uow.session.add(
                DocumentoORM(
                    carregamento_id=int(critico.id),
                    usuario_id=usuario_id,
                    tipo="MINUTA",
                    caminho_arquivo="carregamentos/99/minuta.pdf",
                    nome_arquivo="minuta.pdf",
                    hash_sha256="d" * 64,
                )
            )

        gestao, simulacao, execucao = _build_services(pdf_dir)
        service = RetencaoAutomaticaService(
            gestao_dados_service=gestao,
            simulacao_service=simulacao,
            execucao_service=execucao,
        )
        resultado = service.executar()

        assert resultado.executado is True
        assert resultado.pacotes_removidos == 1

        with UnitOfWork() as uow:
            critico_row = uow.session.scalars(
                select(CarregamentoORM).where(CarregamentoORM.numero_carregamento == "000099")
            ).first()
        assert critico_row is not None

        get_engine().dispose()
        _reset_db_module()


def test_startup_retention_executa_apenas_uma_vez() -> None:
    reset_startup_retention_flag()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        pdf_dir = tmp_path / "documentos"
        _seed_retencao_dataset(tmp_path / "startup_once.db", pdf_dir)
        _configure_bootstrap(tmp_path, "startup_once.db")

        primeiro = run_startup_retention_once()
        segundo = run_startup_retention_once()

        assert primeiro is not None
        assert segundo is None

        get_engine().dispose()
        _reset_db_module()
        reset_startup_retention_flag()


def test_startup_retention_nao_bloqueia_em_falha() -> None:
    reset_startup_retention_flag()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        _configure_bootstrap(tmp_path, "startup_falha.db")

        with patch(
            "carregamentos.services.retencao_automatica_service.RetencaoAutomaticaService.executar",
            side_effect=RuntimeError("falha simulada"),
        ):
            resultado = run_startup_retention_once()

        assert resultado is None

        get_engine().dispose()
        _reset_db_module()
        reset_startup_retention_flag()


def test_retencao_automatica_atualiza_indicadores_gestao_dados() -> None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        pdf_dir = tmp_path / "documentos"
        _seed_retencao_dataset(tmp_path / "auto_painel.db", pdf_dir)
        _configure_bootstrap(tmp_path, "auto_painel.db")

        gestao, simulacao, execucao = _build_services(pdf_dir)
        antes = gestao.obter_painel()
        assert antes.pacote.carregamentos == 1

        service = RetencaoAutomaticaService(
            gestao_dados_service=gestao,
            simulacao_service=simulacao,
            execucao_service=execucao,
        )
        resultado = service.executar()

        assert resultado.painel_atualizado is not None
        assert resultado.painel_atualizado.pacote.carregamentos == 0
        assert resultado.painel_atualizado.uso_banco.bytes_ocupados is not None

        get_engine().dispose()
        _reset_db_module()
