from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal


@dataclass
class NotaFiscalRecord:
    id: int
    chave_nfe: str
    numero_nf: str
    destinatario: str
    status_nf: str
    valor_total: Decimal
    peso_total: Decimal
    volume_total: Decimal
    emitente: str | None = None
    municipio: str | None = None
    uf: str | None = None
    rota: str | None = None
    tipo_xml: str | None = None
    data_emissao: date | None = None
    data_referencia: datetime | None = None
    arquivo_origem: str | None = None
    destinatario_id: int | None = None
    rota_id: int | None = None


@dataclass
class ItemNotaFiscalRecord:
    id: int
    nota_fiscal_id: int
    sequencia: int
    codigo_produto: str
    descricao: str
    quantidade: Decimal
    peso: Decimal
    unidade: str | None = None


class NotaFiscalRepository(ABC):
    @abstractmethod
    def get_by_id(self, nota_fiscal_id: int) -> NotaFiscalRecord | None:
        raise NotImplementedError

    @abstractmethod
    def get_by_chave(self, chave_nfe: str) -> NotaFiscalRecord | None:
        raise NotImplementedError

    @abstractmethod
    def list_all(self) -> list[NotaFiscalRecord]:
        raise NotImplementedError

    @abstractmethod
    def save(self, nota_fiscal: NotaFiscalRecord) -> NotaFiscalRecord:
        raise NotImplementedError

    @abstractmethod
    def list_itens(self, nota_fiscal_id: int) -> list[ItemNotaFiscalRecord]:
        raise NotImplementedError

    @abstractmethod
    def save_itens(self, nota_fiscal_id: int, itens: list[ItemNotaFiscalRecord]) -> list[ItemNotaFiscalRecord]:
        raise NotImplementedError
