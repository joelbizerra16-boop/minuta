from __future__ import annotations

import hashlib
import logging
import re
import time
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import select

from infrastructure.database import get_xml_storage_dir
from infrastructure.models.documento_xml import DocumentoXmlORM
from infrastructure.repositories.documento_xml_repository import DocumentoXmlRecord
from infrastructure.repositories.sql.documento_xml_repository import SqlDocumentoXmlRepository
from infrastructure.storage.xml_chave_extractor import extract_chave_nfe_from_xml_bytes, extract_numero_nf_from_chave
from infrastructure.unit_of_work import UnitOfWork

logger = logging.getLogger("minuta.documento_xml")


@dataclass(frozen=True)
class XmlDocumentalItem:
    file_bytes: bytes
    hash_sha256: str
    original_filename: str
    chave_nfe: str | None = None


@dataclass(frozen=True)
class PersistXmlResult:
    status: str
    chave_nfe: str
    hash_sha256: str


@dataclass(frozen=True)
class XmlBatchPersistResult:
    saved: int
    reused: int
    skipped: int
    failures: int
    elapsed_ms: float
    results: tuple[PersistXmlResult, ...]
    issues: tuple[str, ...]


class DocumentoXmlService:
    def __init__(self, storage_dir: Path | None = None) -> None:
        self._storage_dir = storage_dir or get_xml_storage_dir()

    def persist_raw_xml(
        self,
        file_bytes: bytes,
        *,
        original_filename: str,
        usuario_id: int | None = None,
        hash_sha256: str | None = None,
        chave_nfe: str | None = None,
    ) -> PersistXmlResult | None:
        file_hash = str(hash_sha256 or "").strip() or hashlib.sha256(file_bytes).hexdigest()
        batch_result = self.persist_raw_xml_batch(
            [
                XmlDocumentalItem(
                    file_bytes=file_bytes,
                    hash_sha256=file_hash,
                    original_filename=original_filename,
                    chave_nfe=chave_nfe,
                )
            ],
            usuario_id=usuario_id,
        )
        return batch_result.results[0] if batch_result.results else None

    def persist_raw_xml_batch(
        self,
        items: list[XmlDocumentalItem],
        *,
        usuario_id: int | None = None,
    ) -> XmlBatchPersistResult:
        started = time.perf_counter()
        if not items:
            return XmlBatchPersistResult(0, 0, 0, 0, 0.0, (), ())

        prepared: list[tuple[XmlDocumentalItem, str, str, str, str]] = []
        seen_chaves: set[str] = set()
        skipped = 0
        issues: list[str] = []

        for item in items:
            if not item.file_bytes:
                skipped += 1
                continue
            chave_nfe = str(item.chave_nfe or "").strip() or extract_chave_nfe_from_xml_bytes(item.file_bytes)
            if not chave_nfe:
                skipped += 1
                issues.append(
                    f"XML documental nao persistido (chave ausente): {item.original_filename}"
                )
                logger.warning("XML documental nao persistido: chave ausente em %s", item.original_filename)
                continue
            if chave_nfe in seen_chaves:
                skipped += 1
                continue
            seen_chaves.add(chave_nfe)

            file_hash = str(item.hash_sha256 or "").strip() or hashlib.sha256(item.file_bytes).hexdigest()
            numero_nf = extract_numero_nf_from_chave(chave_nfe) or _fallback_nf_from_filename(item.original_filename)
            safe_name = _sanitize_export_filename(item.original_filename, chave_nfe, numero_nf)
            prepared.append((item, chave_nfe, file_hash, numero_nf, safe_name))

        results: list[PersistXmlResult] = []
        saved = 0
        reused = 0
        failures = 0
        disk_writes = 0
        disk_available = True

        try:
            self._storage_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            disk_available = False
            issues.append(f"Copia local de XML indisponivel; persistencia no banco mantida: {exc}")
            logger.warning("Diretorio local de XML indisponivel %s: %s", self._storage_dir, exc)

        try:
            with UnitOfWork() as uow:
                repo = SqlDocumentoXmlRepository(uow.session)
                chaves = [chave for _, chave, _, _, _ in prepared]
                existing_by_chave = repo.list_by_chaves(chaves)

                for item, chave_nfe, file_hash, numero_nf, safe_name in prepared:
                    relative_path = f"xml_storage/{chave_nfe}.xml"
                    existing = existing_by_chave.get(chave_nfe)
                    if existing is not None and existing.hash_sha256 == file_hash:
                        if not existing.conteudo_xml:
                            saved_record = repo.save(
                                replace(
                                    existing,
                                    tamanho=len(item.file_bytes),
                                    conteudo_xml=bytes(item.file_bytes),
                                )
                            )
                            existing_by_chave[chave_nfe] = saved_record
                            saved += 1
                            results.append(
                                PersistXmlResult(status="saved", chave_nfe=chave_nfe, hash_sha256=file_hash)
                            )
                            logger.info(
                                "XML legado preenchido no banco chave=%s hash=%s",
                                chave_nfe,
                                file_hash[:12],
                            )
                            continue
                        reused += 1
                        results.append(
                            PersistXmlResult(status="reused", chave_nfe=chave_nfe, hash_sha256=file_hash)
                        )
                        logger.info(
                            "XML reutilizado chave=%s hash=%s arquivo=%s",
                            chave_nfe,
                            file_hash[:12],
                            existing.nome_arquivo,
                        )
                        continue

                    if disk_available:
                        absolute_path = self._storage_dir / f"{chave_nfe}.xml"
                        try:
                            if not absolute_path.is_file() or existing is None or existing.hash_sha256 != file_hash:
                                absolute_path.write_bytes(item.file_bytes)
                                disk_writes += 1
                        except OSError as exc:
                            disk_available = False
                            issues.append(
                                f"XML salvo no banco, mas a copia local falhou ({item.original_filename}): {exc}"
                            )
                            logger.warning("Falha na copia local do XML %s: %s", item.original_filename, exc)

                    record = DocumentoXmlRecord(
                        id=int(existing.id) if existing else 0,
                        chave_nfe=chave_nfe,
                        numero_nf=numero_nf,
                        nome_arquivo=safe_name,
                        caminho_arquivo=relative_path,
                        hash_sha256=file_hash,
                        tamanho=len(item.file_bytes),
                        usuario_id=usuario_id,
                        data_importacao=datetime.now(timezone.utc),
                        ativo=True,
                        conteudo_xml=bytes(item.file_bytes),
                    )
                    saved_record = repo.save(record)
                    existing_by_chave[chave_nfe] = saved_record
                    saved += 1
                    results.append(
                        PersistXmlResult(status="saved", chave_nfe=chave_nfe, hash_sha256=file_hash)
                    )
                    logger.info(
                        "XML salvo chave=%s hash=%s tamanho=%s bytes arquivo=%s",
                        chave_nfe,
                        file_hash[:12],
                        len(item.file_bytes),
                        safe_name,
                    )

                if saved > 0:
                    uow.session.flush()
        except Exception as exc:
            failures = max(failures, len(prepared) - reused - saved)
            issues.append(f"Falha na persistencia documental dos XMLs: {exc}")
            logger.exception("Falha na persistencia documental em lote: %s", exc)

        elapsed_ms = (time.perf_counter() - started) * 1000
        logger.info(
            "Persistencia documental XML lote saved=%s reused=%s skipped=%s failures=%s disk_writes=%s tempo_ms=%.1f",
            saved,
            reused,
            skipped,
            failures,
            disk_writes,
            elapsed_ms,
        )
        return XmlBatchPersistResult(
            saved=saved,
            reused=reused,
            skipped=skipped,
            failures=failures,
            elapsed_ms=elapsed_ms,
            results=tuple(results),
            issues=tuple(issues),
        )

    def read_xml_bytes(self, record: DocumentoXmlRecord) -> bytes:
        if record.conteudo_xml:
            return bytes(record.conteudo_xml)

        candidates = [
            self._storage_dir / f"{record.chave_nfe}.xml",
            self._storage_dir / Path(record.caminho_arquivo).name,
        ]
        for candidate in candidates:
            if candidate.is_file():
                payload = candidate.read_bytes()
                self._backfill_database_content(record, payload)
                return payload
        return b""

    @staticmethod
    def _backfill_database_content(record: DocumentoXmlRecord, payload: bytes) -> None:
        if not payload or record.conteudo_xml:
            return
        try:
            with UnitOfWork() as uow:
                repo = SqlDocumentoXmlRepository(uow.session)
                repo.save(
                    replace(
                        record,
                        tamanho=len(payload),
                        conteudo_xml=bytes(payload),
                    )
                )
                uow.session.flush()
        except Exception as exc:
            logger.warning(
                "Falha ao preencher XML legado no banco chave=%s: %s",
                record.chave_nfe,
                exc,
            )


def _sanitize_export_filename(original_filename: str, chave_nfe: str, numero_nf: str) -> str:
    name = Path(str(original_filename or "").strip()).name
    if name.lower().endswith(".xml") and name:
        return name
    if numero_nf:
        return f"NF{numero_nf}.xml"
    return f"{chave_nfe}.xml"


def _fallback_nf_from_filename(filename: str) -> str:
    stem = Path(str(filename or "")).stem
    digits = re.sub(r"\D", "", stem)
    return digits.lstrip("0") or digits or "--"
