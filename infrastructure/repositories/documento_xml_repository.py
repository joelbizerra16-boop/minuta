from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Protocol


@dataclass(frozen=True)
class DocumentoXmlRecord:
    id: int
    chave_nfe: str
    numero_nf: str
    nome_arquivo: str
    caminho_arquivo: str
    hash_sha256: str
    tamanho: int
    usuario_id: int | None
    data_importacao: datetime
    ativo: bool
    conteudo_xml: bytes | None = None


class DocumentoXmlRepository(Protocol):
    def get_by_chave(self, chave_nfe: str) -> DocumentoXmlRecord | None: ...

    def get_by_numero_nf(self, numero_nf: str) -> DocumentoXmlRecord | None: ...

    def list_by_chaves(self, chaves: list[str]) -> dict[str, DocumentoXmlRecord]: ...

    def list_by_numeros_nf(self, numeros_nf: list[str]) -> dict[str, DocumentoXmlRecord]: ...

    def save(self, record: DocumentoXmlRecord) -> DocumentoXmlRecord: ...
