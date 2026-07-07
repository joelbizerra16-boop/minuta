from __future__ import annotations

import tempfile
from pathlib import Path
from unittest.mock import patch

from sqlalchemy import func, select

from carregamentos.models.execucao_retencao import ConfirmacaoRetencao
from carregamentos.repository.simulacao_retencao_repository import ArvoreCarregamentoRaw
from carregamentos.repository.sql_execucao_retencao_repository import SqlExecucaoRetencaoRepository
from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService, RetencaoExecucaoError
from carregamentos.services.gestao_dados_service import GestaoDadosService
from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService
from infrastructure.database import get_engine
from infrastructure.models.carregamento import CarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.models.nota_fiscal import NotaFiscalORM
from infrastructure.models.constants import AUDIT_EVENTO_RETENCAO_DADOS
from infrastructure.unit_of_work import UnitOfWork
from tests.test_gestao_retencao import _seed_retencao_dataset


def _build_services(pdf_dir: Path) -> tuple[GestaoDadosService, SimulacaoRetencaoService, ExecucaoRetencaoService]:
    xml_dir = pdf_dir / "xml"
    gestao = GestaoDadosService(pdf_storage_dir=pdf_dir)
    simulacao = SimulacaoRetencaoService(
        gestao_dados_service=gestao,
        pdf_storage_dir=pdf_dir,
        xml_storage_dir=xml_dir,
    )
    execucao = ExecucaoRetencaoService(
        simulacao_service=simulacao,
        pdf_storage_dir=pdf_dir,
        xml_storage_dir=xml_dir,
    )
    return gestao, simulacao, execucao


def test_executar_retencao_remove_apenas_elegivel() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "retencao_r2.db"
        pdf_dir = tmp_path / "documentos"
        xml_dir = pdf_dir / "xml"
        _seed_retencao_dataset(db_path, pdf_dir)

        xml_file = xml_dir / "xml" / "3010.xml"
        xml_file.parent.mkdir(parents=True, exist_ok=True)
        xml_file.write_bytes(b"<nfe/>")

        _, simulacao, execucao = _build_services(pdf_dir)
        relatorio, confirmacao = execucao.preparar_execucao()

        assert confirmacao.carregamentos == 1
        assert confirmacao.documentos_pdf == 1

        resultado = execucao.executar_retencao(
            confirmacao,
            usuario_id=1,
            usuario_nome="Administrador",
        )
        assert resultado.sucesso is True
        assert resultado.carregamentos_removidos == 1
        assert resultado.revertido is False
        assert resultado.arquivos_pdf_removidos == 1

        with UnitOfWork() as uow:
            session = uow.session
            assert session.scalar(select(func.count()).select_from(CarregamentoORM)) == 1
            assert session.scalar(select(func.count()).select_from(DocumentoORM)) == 1
            assert session.scalar(select(func.count()).select_from(HistoricoOperacionalORM)) == 1
            assert session.scalar(select(func.count()).select_from(EventoAuditoriaORM)) == 2
            assert session.scalar(select(func.count()).select_from(NotaFiscalORM)) == 1
            assert session.scalar(select(func.count()).select_from(DocumentoXmlORM)) == 1
            audit = session.scalars(
                select(EventoAuditoriaORM).where(EventoAuditoriaORM.evento == AUDIT_EVENTO_RETENCAO_DADOS)
            ).all()
            assert len(audit) == 1

        relatorio_pos = simulacao.executar_simulacao()
        assert relatorio_pos.resumo.pacotes_elegiveis == 0

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


def test_executar_retencao_rollback_quando_falha_no_meio() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "retencao_rollback.db"
        pdf_dir = tmp_path / "documentos"
        xml_dir = pdf_dir / "xml"
        _seed_retencao_dataset(db_path, pdf_dir)
        (xml_dir / "xml").mkdir(parents=True, exist_ok=True)
        (xml_dir / "xml" / "3010.xml").write_bytes(b"<nfe/>")

        _, _, execucao = _build_services(pdf_dir)
        _, confirmacao = execucao.preparar_execucao()
        pdf_restante = next(iter(pdf_dir.rglob("*.pdf")), None)
        assert pdf_restante is not None and pdf_restante.is_file()

        original = SqlExecucaoRetencaoRepository.excluir_arvore_carregamento
        chamadas = {"count": 0}

        def falha_no_meio(self, session, arvore: ArvoreCarregamentoRaw) -> None:
            chamadas["count"] += 1
            raise RetencaoExecucaoError("Falha simulada para rollback.")

        with patch.object(SqlExecucaoRetencaoRepository, "excluir_arvore_carregamento", falha_no_meio):
            resultado = execucao.executar_retencao(
                confirmacao,
                usuario_id=1,
                usuario_nome="Administrador",
            )

        assert resultado.sucesso is False
        assert resultado.revertido is True
        assert chamadas["count"] == 1
        assert pdf_restante.is_file()

        with UnitOfWork() as uow:
            session = uow.session
            assert session.scalar(select(func.count()).select_from(CarregamentoORM)) == 2
            assert session.scalar(select(func.count()).select_from(DocumentoORM)) == 2
            audit = session.scalars(
                select(EventoAuditoriaORM).where(EventoAuditoriaORM.evento == AUDIT_EVENTO_RETENCAO_DADOS)
            ).all()
            assert len(audit) == 0

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


def test_preparar_execucao_exige_pacotes_aptos() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "retencao_preparar.db"
        pdf_dir = tmp_path / "documentos"
        _seed_retencao_dataset(db_path, pdf_dir)

        _, simulacao, execucao = _build_services(pdf_dir)
        relatorio, confirmacao = execucao.preparar_execucao()
        assert isinstance(confirmacao, ConfirmacaoRetencao)
        assert confirmacao.carregamento_ids

        confirmacao_invalida = ConfirmacaoRetencao(
            carregamentos=confirmacao.carregamentos,
            notas_fiscais=confirmacao.notas_fiscais,
            documentos_xml=confirmacao.documentos_xml,
            documentos_pdf=confirmacao.documentos_pdf,
            eventos=confirmacao.eventos,
            historicos=confirmacao.historicos,
            espaco_estimado_bytes=confirmacao.espaco_estimado_bytes,
            data_corte=relatorio.data_corte,
            carregamento_ids=(999999,),
        )
        resultado = execucao.executar_retencao(
            confirmacao_invalida,
            usuario_id=1,
            usuario_nome="Administrador",
        )
        assert resultado.sucesso is False
        assert resultado.revertido is True

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


if __name__ == "__main__":
    test_executar_retencao_remove_apenas_elegivel()
    test_executar_retencao_rollback_quando_falha_no_meio()
    test_preparar_execucao_exige_pacotes_aptos()
    print("test_execucao_retencao OK")
