from __future__ import annotations

import tempfile
from pathlib import Path

from carregamentos.services.gestao_dados_service import GestaoDadosService
from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService
from carregamentos.models.simulacao_retencao import SaudePacote
from infrastructure.database import get_engine
from tests.test_gestao_retencao import _seed_retencao_dataset


def test_simulacao_retencao_executavel_sem_destruicao() -> None:
    import infrastructure.database as db_module

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        db_path = tmp_path / "simulacao.db"
        pdf_dir = tmp_path / "documentos"
        xml_dir = pdf_dir / "xml"
        _seed_retencao_dataset(db_path, pdf_dir)

        xml_file = xml_dir / "xml" / "3010.xml"
        xml_file.parent.mkdir(parents=True, exist_ok=True)
        xml_file.write_bytes(b"<nfe/>")

        gestao = GestaoDadosService(pdf_storage_dir=pdf_dir)
        service = SimulacaoRetencaoService(
            gestao_dados_service=gestao,
            pdf_storage_dir=pdf_dir,
            xml_storage_dir=xml_dir,
        )

        relatorio = service.executar_simulacao()
        assert relatorio.simulacao is True
        assert relatorio.resumo.pacotes_elegiveis == 1
        assert relatorio.registros_analisados > 0
        assert relatorio.arquivos_pdf == 1
        assert relatorio.arquivos_xml == 1
        assert relatorio.duracao_ms >= 0
        assert len(relatorio.pacotes) == 1

        pacote = relatorio.pacotes[0]
        assert pacote.itens_carregamento == 1
        assert pacote.notas_fiscais >= 1
        assert pacote.documentos_pdf == 1
        assert pacote.documentos_xml == 1
        assert pacote.historicos == 1
        assert pacote.eventos == 1
        assert pacote.apto_retencao is True
        assert pacote.saude in (SaudePacote.SAUDAVEL, SaudePacote.ATENCAO)
        assert pacote.espaco_recuperavel_bytes > 0

        assert len(relatorio.pacotes_apto_futura_retencao) == 1
        assert len(relatorio.pacotes_requerem_correcao) == 0

        pacote_individual = service.obter_pacote_por_carregamento(pacote.carregamento_id)
        assert pacote_individual is not None
        assert pacote_individual.carregamento_id == pacote.carregamento_id

        get_engine().dispose()
        db_module._engine = None
        db_module._session_factory = None
        db_module._data_root = None
        db_module._pdf_storage_dir = None
        db_module._xml_storage_dir = None


if __name__ == "__main__":
    test_simulacao_retencao_executavel_sem_destruicao()
    print("test_simulacao_retencao OK")
