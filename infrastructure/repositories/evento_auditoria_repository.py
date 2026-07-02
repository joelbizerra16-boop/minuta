from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime


@dataclass
class EventoAuditoriaRecord:
    id: int
    categoria: str
    evento: str
    usuario_id: int | None = None
    entidade_tipo: str | None = None
    entidade_id: int | None = None
    descricao: str | None = None
    metadados_json: str | None = None
    ip_origem: str | None = None
    criado_em: datetime | None = None


class EventoAuditoriaRepository(ABC):
    """Contrato preparado para fases posteriores — sem implementacao ativa na M0.5."""

    @abstractmethod
    def append(self, evento: EventoAuditoriaRecord) -> EventoAuditoriaRecord:
        raise NotImplementedError

    @abstractmethod
    def list_by_entidade(self, entidade_tipo: str, entidade_id: int) -> list[EventoAuditoriaRecord]:
        raise NotImplementedError

    @abstractmethod
    def list_by_usuario(self, usuario_id: int, limit: int = 100) -> list[EventoAuditoriaRecord]:
        raise NotImplementedError
