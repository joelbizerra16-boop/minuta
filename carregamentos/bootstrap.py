from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import TYPE_CHECKING

from carregamentos.repository.carregamento_repository import CarregamentoRepository
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from carregamentos.services.analise_operacional_service import AnaliseOperacionalService
from carregamentos.services.carregamento_service import CarregamentoService
from carregamentos.services.fechamento_service import FechamentoCarregamentoService
from carregamentos.services.historico_carregamento_service import HistoricoCarregamentoService
from carregamentos.services.rastreabilidade_nf_service import RastreabilidadeNfService
from carregamentos.services.xml_export_service import XmlExportService
from infrastructure.database import get_pdf_storage_dir, get_xml_storage_dir

if TYPE_CHECKING:
    from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService
    from carregamentos.services.gestao_capacidade_service import GestaoCapacidadeService
    from carregamentos.services.gestao_dados_service import GestaoDadosService
    from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService

_repository: CarregamentoRepository | None = None
_fechamento_service: FechamentoCarregamentoService | None = None
_analise_operacional_service: AnaliseOperacionalService | None = None
_historico_carregamento_service: HistoricoCarregamentoService | None = None
_gestao_dados_service: GestaoDadosService | None = None
_gestao_capacidade_service: GestaoCapacidadeService | None = None
_simulacao_retencao_service: SimulacaoRetencaoService | None = None
_execucao_retencao_service: ExecucaoRetencaoService | None = None


def configure_carregamentos_storage(data_dir: Path) -> CarregamentoRepository:
    global _repository, _fechamento_service, _analise_operacional_service, _historico_carregamento_service, _gestao_dados_service, _gestao_capacidade_service, _simulacao_retencao_service, _execucao_retencao_service
    from carregamentos.services.execucao_retencao_service import ExecucaoRetencaoService
    from carregamentos.services.gestao_capacidade_service import GestaoCapacidadeService
    from carregamentos.services.gestao_dados_service import GestaoDadosService
    from carregamentos.services.simulacao_retencao_service import SimulacaoRetencaoService

    _ = data_dir
    _repository = SqlCarregamentoRepository()
    _repository.storage_dir.mkdir(parents=True, exist_ok=True)
    _fechamento_service = FechamentoCarregamentoService(_repository, get_pdf_storage_dir())
    _analise_operacional_service = AnaliseOperacionalService(_repository)
    _historico_carregamento_service = HistoricoCarregamentoService(_repository)
    _gestao_dados_service = GestaoDadosService(pdf_storage_dir=get_pdf_storage_dir())
    _simulacao_retencao_service = SimulacaoRetencaoService(
        gestao_dados_service=_gestao_dados_service,
        pdf_storage_dir=get_pdf_storage_dir(),
        xml_storage_dir=get_xml_storage_dir(),
    )
    _execucao_retencao_service = ExecucaoRetencaoService(
        simulacao_service=_simulacao_retencao_service,
        pdf_storage_dir=get_pdf_storage_dir(),
        xml_storage_dir=get_xml_storage_dir(),
    )
    _gestao_capacidade_service = GestaoCapacidadeService(
        gestao_dados_service=_gestao_dados_service,
        simulacao_service=_simulacao_retencao_service,
        execucao_service=_execucao_retencao_service,
    )
    _gestao_dados_service._capacidade_service = _gestao_capacidade_service
    return _repository


def get_carregamento_repository() -> CarregamentoRepository:
    if _repository is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _repository


@lru_cache(maxsize=1)
def get_carregamento_service() -> CarregamentoService:
    return CarregamentoService(get_carregamento_repository(), get_pdf_storage_dir())


@lru_cache(maxsize=1)
def get_xml_export_service() -> XmlExportService:
    return XmlExportService()


@lru_cache(maxsize=1)
def get_rastreabilidade_nf_service() -> RastreabilidadeNfService:
    return RastreabilidadeNfService()


def get_fechamento_service() -> FechamentoCarregamentoService:
    if _fechamento_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _fechamento_service


def get_analise_operacional_service() -> AnaliseOperacionalService:
    if _analise_operacional_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _analise_operacional_service


def invalidate_analise_operacional_cache() -> None:
    if _analise_operacional_service is not None:
        _analise_operacional_service.invalidar_cache()


def get_historico_carregamento_service() -> HistoricoCarregamentoService:
    if _historico_carregamento_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _historico_carregamento_service


def get_gestao_dados_service() -> GestaoDadosService:
    if _gestao_dados_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _gestao_dados_service


def get_gestao_retencao_service() -> GestaoDadosService:
    return get_gestao_dados_service()


def get_gestao_capacidade_service() -> GestaoCapacidadeService:
    if _gestao_capacidade_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _gestao_capacidade_service


def get_simulacao_retencao_service() -> SimulacaoRetencaoService:
    if _simulacao_retencao_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _simulacao_retencao_service


def get_execucao_retencao_service() -> ExecucaoRetencaoService:
    if _execucao_retencao_service is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _execucao_retencao_service
