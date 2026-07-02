from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime


@dataclass
class ConfiguracaoRecord:
    id: int
    chave: str
    valor: str
    categoria: str = "GERAL"
    tipo_valor: str = "STRING"
    descricao: str | None = None
    atualizado_por_usuario_id: int | None = None
    atualizado_em: datetime | None = None


class ConfiguracaoRepository(ABC):
    @abstractmethod
    def get_by_chave(self, chave: str) -> ConfiguracaoRecord | None:
        raise NotImplementedError

    @abstractmethod
    def list_all(self) -> list[ConfiguracaoRecord]:
        raise NotImplementedError

    @abstractmethod
    def save(self, configuracao: ConfiguracaoRecord) -> ConfiguracaoRecord:
        raise NotImplementedError

    @abstractmethod
    def delete(self, chave: str) -> bool:
        raise NotImplementedError
