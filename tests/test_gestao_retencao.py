from __future__ import annotations

import tempfile
from datetime import date, time, timedelta
from pathlib import Path

from carregamentos.repository.sql_retencao_repository import SqlRetencaoRepository
from carregamentos.services.gestao_dados_service import GestaoDadosService
from core.retention_policy import RETENTION_DAYS, DATABASE_STORAGE_LIMIT_BYTES, retention_days_before_today
from infrastructure.database import configure_database, get_engine
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.models.perfil import PerfilORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.unit_of_work import UnitOfWork


def _seed_retencao_dataset(db_path: Path, pdf_dir: Path) -> None:
    from infrastructure.schema import ensure_full_schema

    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=db_path.parent,
        pdf_storage_dir=pdf_dir,
        xml_storage_dir=pdf_dir / "xml",
    )
    ensure_full_schema()

    hoje = date.today()
    data_elegivel = hoje - timedelta(days=retention_days_before_today() + 3)
    data_mantida = hoje - timedelta(days=1)

    with UnitOfWork() as uow:
        uow.session.add(PerfilORM(id=2, codigo="OPERADOR", nome="Operador"))
        uow.session.flush()
        usuario = UsuarioORM(
            uuid="00000000-0000-4000-8000-000000000001",
            nome="Teste",
            usuario="teste",
            senha_hash="pbkdf2_sha256$100000$aaaaaaaaaaaaaaaa$bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            perfil_id=2,
            perfil="OPERADOR",
        )
        uow.session.add(usuario)
        uow.session.flush()

        nf = NotaFiscalORM(
            chave_nfe="35260600000000000000000000000000000000000040",
            numero_nf="3010",
            destinatario="Cliente Teste",
            status_nf="AUTORIZADA",
        )
        uow.session.add(nf)
        uow.session.flush()
        uow.session.add(
            ItemNotaFiscalORM(
                nota_fiscal_id=int(nf.id),
                sequencia=1,
                codigo_produto="P001",
                descricao="Produto",
                quantidade=1,
                peso=1,
            )
        )

        uow.session.add(
            DocumentoXmlORM(
                chave_nfe=nf.chave_nfe,
                numero_nf=nf.numero_nf,
                nome_arquivo="3010.xml",
                caminho_arquivo="xml/3010.xml",
                hash_sha256="a" * 64,
                tamanho=2048,
                usuario_id=int(usuario.id),
            )
        )

        for index, data_carregamento in enumerate((data_elegivel, data_mantida), start=1):
            carregamento = CarregamentoORM(
                numero_carregamento=f"{index:06d}",
                usuario_id=int(usuario.id),
                data=data_carregamento,
                hora=time(10, 0),
                motorista="Motorista",
                placa="ABC1D23",
                modalidade="VEICULO",
                status="FINALIZADO",
            )
            uow.session.add(carregamento)
            uow.session.flush()
            uow.session.add(
                ItemCarregamentoORM(
                    carregamento_id=int(carregamento.id),
                    nota_fiscal_id=int(nf.id),
                    chave_nfe=nf.chave_nfe,
                    numero_nf=nf.numero_nf,
                    codigo_produto="P001",
                    descricao="Produto",
                    sequencia=1,
                )
            )
            uow.session.add(
                HistoricoOperacionalORM(
                    carregamento_id=int(carregamento.id),
                    usuario_id=int(usuario.id),
                    evento="FINALIZACAO",
                    descricao="Teste",
                )
            )
            uow.session.add(
                EventoAuditoriaORM(
                    usuario_id=int(usuario.id),
                    categoria="CARREGAMENTO",
                    evento="FINALIZACAO_CARREGAMENTO",
                    entidade_tipo="carregamento",
                    entidade_id=int(carregamento.id),
                    descricao="Teste",
                )
            )
            relative_pdf = f"carregamentos/{carregamento.id}/minuta_carregamento.pdf"
            absolute_pdf = pdf_dir / relative_pdf
            absolute_pdf.parent.mkdir(parents=True, exist_ok=True)
            absolute_pdf.write_bytes(b"%PDF-1.4 test")
            uow.session.add(
                DocumentoORM(
                    carregamento_id=int(carregamento.id),
                    usuario_id=int(usuario.id),
                    tipo="MINUTA",
                    caminho_arquivo=relative_pdf,
                    nome_arquivo="minuta_carregamento.pdf",
                    hash_sha256="c" * 64,
                )
            )

    get_engine().dispose()


def test_gestao_dados_painel_e_pacote_retencao() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "retencao.db"
        pdf_dir = tmp_path / "documentos"
        _seed_retencao_dataset(db_path, pdf_dir)

        service = GestaoDadosService(pdf_storage_dir=pdf_dir)
        painel = service.obter_painel()
        pacote = painel.pacote

        assert RETENTION_DAYS == 8
        assert painel.politica_dias_mantidos == 8
        assert painel.politica_status == "Ativa"
        assert pacote.carregamentos == 1
        assert pacote.itens_carregamento == 1
        assert pacote.historicos == 1
        assert pacote.eventos == 1
        assert pacote.documentos_pdf == 1
        assert pacote.notas_fiscais == 1
        assert pacote.itens_nota_fiscal == 1
        assert pacote.documentos_xml == 1
        assert pacote.espaco_xmls_bytes == 2048
        assert pacote.espaco_pdfs_bytes > 0
        assert pacote.possui_elegiveis is True
        assert painel.uso_banco.bytes_ocupados is not None
        assert painel.uso_banco.bytes_limite == DATABASE_STORAGE_LIMIT_BYTES
        assert painel.capacidade is not None
        assert painel.simulacao is True

        preview = service.analisar()
        assert preview.carregamentos_elegiveis == pacote.carregamentos
        assert preview.espaco_total_estimado_bytes == pacote.espaco_recuperavel_bytes

        repo = SqlRetencaoRepository()
        assert repo.possui_carregamentos_elegiveis(service.calcular_data_corte()) is True

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


def test_gestao_retencao_compat() -> None:
    test_gestao_dados_painel_e_pacote_retencao()


if __name__ == "__main__":
    test_gestao_dados_painel_e_pacote_retencao()
    print("test_gestao_dados OK")
