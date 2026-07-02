from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime


@dataclass
class HistoricoRecord:
    id: int
    carregamento_id: int
    usuario_id: int
    evento: str
    descricao: str | None = None
    criado_em: datetime | None = None


class HistoricoRepository(ABC):
    @abstractmethod
    def list_by_carregamento(self, carregamento_id: int) -> list[HistoricoRecord]:
        raise NotImplementedError

    @abstractmethod
    def append(self, historico: HistoricoRecord) -> HistoricoRecord:
        raise NotImplementedError
