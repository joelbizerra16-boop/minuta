from __future__ import annotations

import tempfile
from datetime import date, time, timedelta
from pathlib import Path

from carregamentos.repository.sql_retencao_repository import SqlRetencaoRepository
from carregamentos.services.gestao_capacidade_service import (
    GestaoCapacidadeService,
    classificar_faixa_capacidade,
    limite_operacional_bytes,
    montar_barra_capacidade,
    projetar_uso_apos_recuperacao,
)
from carregamentos.services.gestao_dados_service import GestaoDadosService
from carregamentos.models.capacidade import FaixaCapacidade
from core.retention_policy import (
    CAPACITY_ORANGE_MIN_PERCENT,
    CAPACITY_RED_MIN_PERCENT,
    CAPACITY_YELLOW_MIN_PERCENT,
    DATABASE_STORAGE_LIMIT_BYTES,
    retention_days_before_today,
)
from infrastructure.database import configure_database, get_engine
from infrastructure.models.carregamento import CarregamentoORM, ItemCarregamentoORM
from infrastructure.models.documento import DocumentoORM
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.models.evento_auditoria import EventoAuditoriaORM
from infrastructure.models.historico import HistoricoOperacionalORM
from infrastructure.models.nota_fiscal import ItemNotaFiscalORM, NotaFiscalORM
from infrastructure.models.perfil import PerfilORM
from infrastructure.models.usuario import UsuarioORM
from infrastructure.services.database_usage_service import UsoBancoDados
from infrastructure.unit_of_work import UnitOfWork


def _seed_capacidade_dataset(db_path: Path, pdf_dir: Path) -> None:
    from infrastructure.schema import ensure_full_schema

    configure_database(
        database_url=f"sqlite:///{db_path.as_posix()}",
        data_root=db_path.parent,
        pdf_storage_dir=pdf_dir,
        xml_storage_dir=pdf_dir / "xml",
    )
    ensure_full_schema()

    hoje = date.today()
    data_antiga = hoje - timedelta(days=retention_days_before_today() + 5)
    data_intermediaria = hoje - timedelta(days=retention_days_before_today() + 2)

    with UnitOfWork() as uow:
        uow.session.add(PerfilORM(id=2, codigo="OPERADOR", nome="Operador"))
        uow.session.flush()
        usuario = UsuarioORM(
            uuid="00000000-0000-4000-8000-000000000002",
            nome="Teste Cap",
            usuario="testecap",
            senha_hash="pbkdf2_sha256$100000$aaaaaaaaaaaaaaaa$bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            perfil_id=2,
            perfil="OPERADOR",
        )
        uow.session.add(usuario)
        uow.session.flush()

        nf = NotaFiscalORM(
            chave_nfe="35260600000000000000000000000000000000000041",
            numero_nf="4010",
            destinatario="Cliente Cap",
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
                nome_arquivo="4010.xml",
                caminho_arquivo="xml/4010.xml",
                hash_sha256="b" * 64,
                tamanho=1024,
                usuario_id=int(usuario.id),
            )
        )

        for index, data_carregamento in enumerate((data_antiga, data_intermediaria), start=1):
            carregamento = CarregamentoORM(
                numero_carregamento=f"{index:06d}",
                usuario_id=int(usuario.id),
                data=data_carregamento,
                hora=time(9, 0),
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
            relative_pdf = f"carregamentos/{carregamento.id}/minuta.pdf"
            absolute_pdf = pdf_dir / relative_pdf
            absolute_pdf.parent.mkdir(parents=True, exist_ok=True)
            absolute_pdf.write_bytes(b"%PDF-1.4 capacidade")
            uow.session.add(
                DocumentoORM(
                    carregamento_id=int(carregamento.id),
                    usuario_id=int(usuario.id),
                    tipo="MINUTA",
                    caminho_arquivo=relative_pdf,
                    nome_arquivo="minuta.pdf",
                    hash_sha256="d" * 64,
                )
            )

    get_engine().dispose()


def test_limite_operacional_500mb() -> None:
    assert limite_operacional_bytes() == 500 * 1024 * 1024
    assert DATABASE_STORAGE_LIMIT_BYTES == 500 * 1024 * 1024


def test_classificacao_faixas_capacidade() -> None:
    assert classificar_faixa_capacidade(70.0) == FaixaCapacidade.VERDE
    assert classificar_faixa_capacidade(85.0) == FaixaCapacidade.AMARELA
    assert classificar_faixa_capacidade(CAPACITY_YELLOW_MIN_PERCENT) == FaixaCapacidade.AMARELA
    assert classificar_faixa_capacidade(CAPACITY_ORANGE_MIN_PERCENT) == FaixaCapacidade.LARANJA
    assert classificar_faixa_capacidade(94.0) == FaixaCapacidade.LARANJA
    assert classificar_faixa_capacidade(CAPACITY_RED_MIN_PERCENT) == FaixaCapacidade.VERMELHA


def test_barra_capacidade_visual() -> None:
    assert montar_barra_capacidade(0) == "░░░░░░░░░░"
    assert montar_barra_capacidade(34) == "███░░░░░░░"
    assert montar_barra_capacidade(91) == "█████████░"
    assert len(montar_barra_capacidade(100)) == 10


def test_projecao_uso_apos_recuperacao() -> None:
    uso = UsoBancoDados(
        motor="SQLite",
        bytes_ocupados=450_000_000,
        bytes_limite=500_000_000,
        bytes_disponiveis=50_000_000,
        utilizacao_percentual=90.0,
        medicao_direta=False,
        rotulo_ocupacao="teste",
    )
    ocupado_apos, pct_apos = projetar_uso_apos_recuperacao(uso, 50_000_000)
    assert ocupado_apos == 400_000_000
    assert pct_apos == 80.0


def test_previa_dia_mais_antigo() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "capacidade.db"
        pdf_dir = tmp_path / "documentos"
        _seed_capacidade_dataset(db_path, pdf_dir)

        gestao = GestaoDadosService(pdf_storage_dir=pdf_dir)
        service = GestaoCapacidadeService(gestao_dados_service=gestao)

        data_corte = gestao.calcular_data_corte()
        repo = SqlRetencaoRepository()
        data_mais_antiga = repo.obter_data_mais_antiga_elegivel(data_corte)
        assert data_mais_antiga is not None

        previa = service.montar_previa_dia_mais_antigo()
        assert previa is not None
        assert previa.data_alvo == data_mais_antiga
        assert previa.carregamentos == 1
        assert len(previa.carregamento_ids) == 1
        assert previa.percentual_apos is not None
        assert previa.percentual_apos <= (previa.percentual_atual or 100)

        painel = gestao.obter_painel()
        assert painel.capacidade.faixa in FaixaCapacidade
        assert painel.uso_banco.bytes_limite == DATABASE_STORAGE_LIMIT_BYTES

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


if __name__ == "__main__":
    test_limite_operacional_500mb()
    test_classificacao_faixas_capacidade()
    test_barra_capacidade_visual()
    test_projecao_uso_apos_recuperacao()
    test_previa_dia_mais_antigo()
    print("test_gestao_capacidade OK")
