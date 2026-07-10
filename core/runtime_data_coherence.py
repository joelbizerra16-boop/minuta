"""Assinaturas de dados operacionais baseadas no PostgreSQL (fonte oficial de leitura)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from infrastructure.repositories.sql.configuracao_repository import SqlConfiguracaoRepository
from infrastructure.storage.config_storage import (
    CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS,
    CONFIG_CHAVE_LOTES,
    CONFIG_CHAVE_SEPARACAO,
    CONFIG_CHAVE_SEPARACAO_EXCLUIDOS,
)
from infrastructure.storage.xml_storage import SqlXmlRecordRepository

# Cache de processo: evita 3–5 round-trips de assinatura a cada rerun Streamlit.
_SIGNATURE_GENERATION = 0
_SIGNATURE_CACHE: dict[str, tuple[int, Any]] = {}


@dataclass(frozen=True, slots=True)
class DataSignature:
    count: int
    revision: str | None


def invalidate_data_signature_cache() -> None:
    """Invalida assinaturas apos escrita confirmada (ou em testes)."""
    global _SIGNATURE_GENERATION
    _SIGNATURE_GENERATION += 1
    _SIGNATURE_CACHE.clear()


def _cached(key: str, factory):
    generation = _SIGNATURE_GENERATION
    hit = _SIGNATURE_CACHE.get(key)
    if hit is not None and hit[0] == generation:
        return hit[1]
    value = factory()
    _SIGNATURE_CACHE[key] = (generation, value)
    return value


def _format_revision(value: datetime | None) -> str | None:
    if value is None:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    return value.isoformat()


def get_xml_data_signature() -> DataSignature:
    def _build() -> DataSignature:
        repository = SqlXmlRecordRepository()
        return DataSignature(
            count=repository.count_records(),
            revision=_format_revision(repository.get_last_updated_at()),
        )

    return _cached("xml", _build)


def get_config_data_signature(chave: str) -> DataSignature:
    def _build() -> DataSignature:
        record = SqlConfiguracaoRepository().get_by_chave(chave)
        if record is None:
            return DataSignature(count=0, revision=None)
        return DataSignature(
            count=len(str(record.valor or "")),
            revision=_format_revision(record.atualizado_em),
        )

    return _cached(f"config:{chave}", _build)


def get_reference_data_signature() -> tuple[DataSignature, DataSignature]:
    return _cached(
        "reference",
        lambda: (
            get_xml_data_signature(),
            get_config_data_signature(CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS),
        ),
    )


def get_operational_data_signature() -> tuple[
    tuple[DataSignature, DataSignature],
    DataSignature,
    DataSignature,
]:
    return _cached(
        "operational",
        lambda: (
            get_reference_data_signature(),
            get_config_data_signature(CONFIG_CHAVE_SEPARACAO),
            get_config_data_signature(CONFIG_CHAVE_SEPARACAO_EXCLUIDOS),
        ),
    )


def get_classificacao_version_token() -> tuple[int, str | None]:
    signature = get_config_data_signature(CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS)
    return signature.count, signature.revision


def get_config_storage_status(chave: str) -> tuple[bool, str]:
    signature = get_config_data_signature(chave)
    if signature.count == 0 and signature.revision is None:
        return False, ""
    if signature.revision:
        return True, signature.revision
    return True, ""


def get_separacao_storage_status_from_db() -> tuple[bool, str]:
    return get_config_storage_status(CONFIG_CHAVE_SEPARACAO)


def get_lotes_storage_status_from_db() -> tuple[bool, str]:
    return get_config_storage_status(CONFIG_CHAVE_LOTES)


__all__ = [
    "CONFIG_CHAVE_CLASSIFICACAO_PRODUTOS",
    "CONFIG_CHAVE_LOTES",
    "CONFIG_CHAVE_SEPARACAO",
    "CONFIG_CHAVE_SEPARACAO_EXCLUIDOS",
    "DataSignature",
    "get_classificacao_version_token",
    "get_config_data_signature",
    "get_config_storage_status",
    "get_lotes_storage_status_from_db",
    "get_operational_data_signature",
    "get_reference_data_signature",
    "get_separacao_storage_status_from_db",
    "get_xml_data_signature",
    "invalidate_data_signature_cache",
]
