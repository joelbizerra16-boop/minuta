from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from carregamentos.repository.carregamento_repository import CarregamentoRepository
from carregamentos.repository.sql_carregamento_repository import SqlCarregamentoRepository
from carregamentos.services.analise_operacional_service import AnaliseOperacionalService
from carregamentos.services.carregamento_service import CarregamentoService
from carregamentos.services.fechamento_service import FechamentoCarregamentoService
from carregamentos.services.historico_carregamento_service import HistoricoCarregamentoService
from carregamentos.services.rastreabilidade_nf_service import RastreabilidadeNfService
from infrastructure.database import get_pdf_storage_dir

_repository: CarregamentoRepository | None = None
_fechamento_service: FechamentoCarregamentoService | None = None
_analise_operacional_service: AnaliseOperacionalService | None = None
_historico_carregamento_service: HistoricoCarregamentoService | None = None


def configure_carregamentos_storage(data_dir: Path) -> CarregamentoRepository:
    global _repository, _fechamento_service, _analise_operacional_service, _historico_carregamento_service
    _ = data_dir
    _repository = SqlCarregamentoRepository()
    _repository.storage_dir.mkdir(parents=True, exist_ok=True)
    _fechamento_service = FechamentoCarregamentoService(_repository, get_pdf_storage_dir())
    _analise_operacional_service = AnaliseOperacionalService(_repository)
    _historico_carregamento_service = HistoricoCarregamentoService(_repository)
    return _repository


def get_carregamento_repository() -> CarregamentoRepository:
    if _repository is None:
        raise RuntimeError("Carregamentos storage not configured. Call configure_carregamentos_storage first.")
    return _repository


@lru_cache(maxsize=1)
def get_carregamento_service() -> CarregamentoService:
    return CarregamentoService(get_carregamento_repository(), get_pdf_storage_dir())


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
