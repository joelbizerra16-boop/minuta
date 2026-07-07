from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import date

from carregamentos.models.retencao import RetencaoContagensArvore


class RetencaoRepository(ABC):
    @abstractmethod
    def possui_carregamentos_elegiveis(self, data_corte: date) -> bool:
        raise NotImplementedError

    def possui_carregamentos_expirados(self, data_corte: date) -> bool:
        """Alias legado — preferir possui_carregamentos_elegiveis."""
        return self.possui_carregamentos_elegiveis(data_corte)

    @abstractmethod
    def coletar_contagens_arvore(self, data_corte: date) -> RetencaoContagensArvore:
        raise NotImplementedError

    @abstractmethod
    def obter_data_mais_antiga_elegivel(self, data_corte: date) -> date | None:
        raise NotImplementedError

    @abstractmethod
    def listar_carregamento_ids_por_data(self, data_corte: date, data_alvo: date) -> tuple[int, ...]:
        raise NotImplementedError

    @abstractmethod
    def coletar_contagens_arvore_por_data(self, data_corte: date, data_alvo: date) -> RetencaoContagensArvore:
        raise NotImplementedError
