from __future__ import annotations

import json
from typing import Any

from sqlalchemy import select
from sqlalchemy.orm import Session

from infrastructure.models.configuracao import ConfiguracaoORM
from infrastructure.models.constants import CONFIG_TIPO_JSON
from infrastructure.repositories.configuracao_repository import ConfiguracaoRepository, ConfiguracaoRecord
from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
from infrastructure.unit_of_work import UnitOfWork

CONFIG_CHAVE_SEPARACAO = "separacao.records"
CONFIG_CHAVE_LOTES = "lotes.records"
CONFIG_CHAVE_SEPARACAO_EXCLUIDOS = "separacao.excluidos"
CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS = "classificacao.produtos"


class SqlJsonConfigStorage:
    """Persistencia de estruturas legadas em configuracao.valor (tipo JSON)."""

    def __init__(self, repository: ConfiguracaoRepository | None = None) -> None:
        self._repository = repository or SqlConfiguracaoRepository()

    def load_list(self, chave: str, default: list[Any] | None = None) -> list[Any]:
        record = self._repository.get_by_chave(chave)
        if record is None or not str(record.valor or "").strip():
            return list(default or [])
        try:
            payload = json.loads(record.valor)
        except json.JSONDecodeError:
            return list(default or [])
        return payload if isinstance(payload, list) else list(default or [])

    def save_list(self, chave: str, records: list[Any], *, categoria: str = "OPERACIONAL") -> None:
        self._repository.save(
            ConfiguracaoRecord(
                id=0,
                chave=chave,
                valor=json.dumps(records, ensure_ascii=False),
                categoria=categoria,
                tipo_valor=CONFIG_TIPO_JSON,
            )
        )

    def load_set(self, chave: str) -> set[str]:
        record = self._repository.get_by_chave(chave)
        if record is None or not str(record.valor or "").strip():
            return set()
        try:
            payload = json.loads(record.valor)
        except json.JSONDecodeError:
            return set()
        if isinstance(payload, list):
            return {str(item) for item in payload}
        return set()

    def save_set(self, chave: str, values: set[str], *, categoria: str = "OPERACIONAL") -> None:
        self.save_list(chave, sorted(values), categoria=categoria)

    def load_dict(self, chave: str) -> dict[str, Any]:
        record = self._repository.get_by_chave(chave)
        if record is None or not str(record.valor or "").strip():
            return {}
        try:
            payload = json.loads(record.valor)
        except json.JSONDecodeError:
            return {}
        return payload if isinstance(payload, dict) else {}

    def save_dict(self, chave: str, payload: dict[str, Any], *, categoria: str = "CONFIGURACAO") -> None:
        self._repository.save(
            ConfiguracaoRecord(
                id=0,
                chave=chave,
                valor=json.dumps(payload, ensure_ascii=False),
                categoria=categoria,
                tipo_valor=CONFIG_TIPO_JSON,
            )
        )
