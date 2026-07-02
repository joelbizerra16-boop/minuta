from __future__ import annotations

from abc import ABC, abstractmethod

from carregamentos.models.rastreabilidade_nf import RastreabilidadeNfRelatorio


class RastreabilidadeNfRepository(ABC):
    @abstractmethod
    def buscar_por_termo(self, termo: str) -> RastreabilidadeNfRelatorio | None:
        raise NotImplementedError
