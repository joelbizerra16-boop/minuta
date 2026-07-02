from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime


@dataclass
class DocumentoRecord:
    id: int
    carregamento_id: int
    usuario_id: int
    tipo: str
    caminho_arquivo: str
    nome_arquivo: str
    hash_sha256: str
    criado_em: datetime | None = None


class DocumentoRepository(ABC):
    @abstractmethod
    def get_by_id(self, documento_id: int) -> DocumentoRecord | None:
        raise NotImplementedError

    @abstractmethod
    def list_by_carregamento(self, carregamento_id: int) -> list[DocumentoRecord]:
        raise NotImplementedError

    @abstractmethod
    def save(self, documento: DocumentoRecord) -> DocumentoRecord:
        raise NotImplementedError
