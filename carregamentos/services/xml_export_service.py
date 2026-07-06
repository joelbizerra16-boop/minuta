from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from pathlib import Path

from sqlalchemy import select

from carregamentos.models.carregamento import normalize_chave_nfe, normalize_nf_number
from infrastructure.models.carregamento import ItemCarregamentoORM
from infrastructure.repositories.documento_xml_repository import DocumentoXmlRecord
from infrastructure.repositories.sql.documento_xml_repository import SqlDocumentoXmlRepository
from infrastructure.services.documento_xml_service import DocumentoXmlService
from infrastructure.unit_of_work import UnitOfWork

logger = logging.getLogger("minuta.xml_export")


@dataclass(frozen=True)
class XmlExportEntry:
    nome_arquivo: str
    conteudo: bytes
    chave_nfe: str
    numero_nf: str


@dataclass(frozen=True)
class XmlExportResult:
    entries: tuple[XmlExportEntry, ...]
    missing_nfs: tuple[str, ...]
    elapsed_ms: float


class XmlExportService:
    def __init__(
        self,
        *,
        documento_xml_service: DocumentoXmlService | None = None,
        storage_dir: Path | None = None,
    ) -> None:
        self._documento_service = documento_xml_service or DocumentoXmlService(storage_dir=storage_dir)

    def collect_xmls_for_carregamento(self, carregamento_id: int) -> XmlExportResult:
        started = time.perf_counter()
        entries: list[XmlExportEntry] = []
        missing: list[str] = []

        with UnitOfWork() as uow:
            nfs = self._list_nf_references_in_session(uow.session, carregamento_id)
            repo = SqlDocumentoXmlRepository(uow.session)

            chaves = [chave for chave, _ in nfs if chave]
            records_by_chave = repo.list_by_chaves(chaves)

            numeros_fallback: list[str] = []
            for chave_nfe, numero_nf in nfs:
                if chave_nfe and chave_nfe in records_by_chave:
                    continue
                if numero_nf:
                    numeros_fallback.append(numero_nf)
            records_by_numero = repo.list_by_numeros_nf(numeros_fallback)

            seen_chaves: set[str] = set()
            for chave_nfe, numero_nf in nfs:
                if chave_nfe and chave_nfe in seen_chaves:
                    continue
                if chave_nfe:
                    seen_chaves.add(chave_nfe)

                record: DocumentoXmlRecord | None = None
                if chave_nfe:
                    record = records_by_chave.get(chave_nfe)
                if record is None and numero_nf:
                    record = records_by_numero.get(numero_nf)

                label = numero_nf or chave_nfe or "--"
                if record is None:
                    missing.append(label)
                    logger.warning(
                        "XML nao encontrado para exportacao carregamento_id=%s nf=%s chave=%s",
                        carregamento_id,
                        numero_nf,
                        chave_nfe,
                    )
                    continue

                content = self._documento_service.read_xml_bytes(record)
                if not content:
                    missing.append(label)
                    logger.warning(
                        "XML nao encontrado em disco carregamento_id=%s nf=%s chave=%s caminho=%s",
                        carregamento_id,
                        numero_nf,
                        chave_nfe,
                        record.caminho_arquivo,
                    )
                    continue

                entries.append(
                    XmlExportEntry(
                        nome_arquivo=record.nome_arquivo,
                        conteudo=content,
                        chave_nfe=record.chave_nfe,
                        numero_nf=record.numero_nf,
                    )
                )
                logger.info(
                    "XML exportado carregamento_id=%s nf=%s arquivo=%s",
                    carregamento_id,
                    record.numero_nf,
                    record.nome_arquivo,
                )

        elapsed_ms = (time.perf_counter() - started) * 1000
        logger.info(
            "Exportacao XML carregamento_id=%s encontrados=%s ausentes=%s tempo_ms=%.1f",
            carregamento_id,
            len(entries),
            len(missing),
            elapsed_ms,
        )
        return XmlExportResult(entries=tuple(entries), missing_nfs=tuple(missing), elapsed_ms=elapsed_ms)

    @staticmethod
    def _list_nf_references_in_session(session, carregamento_id: int) -> list[tuple[str, str]]:
        rows = session.scalars(
            select(ItemCarregamentoORM)
            .where(ItemCarregamentoORM.carregamento_id == int(carregamento_id))
            .order_by(ItemCarregamentoORM.sequencia, ItemCarregamentoORM.id)
        ).all()

        ordered: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for row in rows:
            chave = normalize_chave_nfe(row.chave_nfe or "")
            numero = normalize_nf_number(row.numero_nf) or str(row.numero_nf or "").strip()
            key = (chave, numero)
            if key in seen:
                continue
            seen.add(key)
            ordered.append(key)
        return ordered
