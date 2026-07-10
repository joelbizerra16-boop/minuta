from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from carregamentos.models.carregamento import Carregamento, CarregamentoFiltro


class CarregamentoRepository(ABC):
    @abstractmethod
    def list_all(self) -> list[Carregamento]:
        raise NotImplementedError

    def list_by_item_identidades(
        self,
        *,
        chaves_nfe: set[str] | None = None,
        numeros_nf: set[str] | None = None,
    ) -> list[Carregamento]:
        """Default: filtra list_all em memoria (implementacoes SQL devem sobrescrever)."""
        chaves = {str(value).strip() for value in (chaves_nfe or set()) if str(value or "").strip()}
        numeros = {str(value).strip() for value in (numeros_nf or set()) if str(value or "").strip()}
        if not chaves and not numeros:
            return []
        from carregamentos.models.carregamento import normalize_chave_nfe, normalize_nf_number

        resultados: list[Carregamento] = []
        for carregamento in self.list_all():
            for item in carregamento.itens:
                item_chave = normalize_chave_nfe(item.chave_nfe)
                item_nf = normalize_nf_number(item.nf)
                if (item_chave and item_chave in chaves) or (item_nf and item_nf in numeros) or (
                    str(item.nf or "").strip() in numeros
                ):
                    resultados.append(carregamento)
                    break
        return resultados

    @abstractmethod
    def get_by_id(self, carregamento_id: int) -> Carregamento | None:
        raise NotImplementedError

    @abstractmethod
    def get_by_numero(self, numero_carregamento: str) -> Carregamento | None:
        raise NotImplementedError

    @abstractmethod
    def save(self, carregamento: Carregamento) -> Carregamento:
        raise NotImplementedError

    @abstractmethod
    def search(self, filtro: CarregamentoFiltro) -> list[Carregamento]:
        raise NotImplementedError

    @abstractmethod
    def proximo_numero_carregamento(self) -> str:
        raise NotImplementedError


class JsonCarregamentoRepository(CarregamentoRepository):
    def __init__(self, index_path: Path, storage_dir: Path):
        self._index_path = index_path
        self._storage_dir = storage_dir

    @property
    def storage_dir(self) -> Path:
        return self._storage_dir

    def list_all(self) -> list[Carregamento]:
        return self._load_carregamentos()

    def get_by_id(self, carregamento_id: int) -> Carregamento | None:
        for carregamento in self._load_carregamentos():
            if carregamento.id == carregamento_id:
                return carregamento
        return None

    def get_by_numero(self, numero_carregamento: str) -> Carregamento | None:
        normalized = str(numero_carregamento or "").strip()
        for carregamento in self._load_carregamentos():
            if carregamento.numero_carregamento == normalized:
                return carregamento
        return None

    def save(self, carregamento: Carregamento) -> Carregamento:
        carregamentos = self._load_carregamentos()
        if carregamento.id <= 0:
            carregamento.id = self._next_id(carregamentos)
        updated = False
        for index, current in enumerate(carregamentos):
            if current.id == carregamento.id:
                carregamentos[index] = carregamento
                updated = True
                break
        if not updated:
            carregamentos.append(carregamento)
        self._write_carregamentos(carregamentos)
        return carregamento

    def search(self, filtro: CarregamentoFiltro) -> list[Carregamento]:
        results = self._load_carregamentos()
        if filtro.data_inicial:
            results = [item for item in results if item.data >= filtro.data_inicial]
        if filtro.data_final:
            results = [item for item in results if item.data <= filtro.data_final]
        return sorted(results, key=lambda item: (item.data, item.hora, item.id), reverse=True)

    def _load_carregamentos(self) -> list[Carregamento]:
        import json

        self._index_path.parent.mkdir(parents=True, exist_ok=True)
        if not self._index_path.is_file():
            return []

        try:
            payload = json.loads(self._index_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            return []

        raw_items = payload.get("carregamentos", []) if isinstance(payload, dict) else []
        return [Carregamento.from_dict(item) for item in raw_items if isinstance(item, dict)]

    def _write_carregamentos(self, carregamentos: list[Carregamento]) -> None:
        import json

        self._index_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "carregamentos": [
                carregamento.to_dict()
                for carregamento in sorted(carregamentos, key=lambda item: item.id)
            ]
        }
        self._index_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _next_id(self, carregamentos: list[Carregamento]) -> int:
        if not carregamentos:
            return 1
        return max(item.id for item in carregamentos) + 1

    def proximo_numero_carregamento(self) -> str:
        import re

        carregamentos = self._load_carregamentos()
        max_seq = 0
        max_id = 0
        for carregamento in carregamentos:
            max_id = max(max_id, int(carregamento.id or 0))
            normalized = str(carregamento.numero_carregamento or "").strip()
            if re.fullmatch(r"\d{6}", normalized):
                max_seq = max(max_seq, int(normalized))
        return f"{max(max_id, max_seq) + 1:06d}"
