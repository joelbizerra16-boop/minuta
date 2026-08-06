from __future__ import annotations

from datetime import datetime

from sqlalchemy import Boolean, CHAR, DateTime, ForeignKey, Index, Integer, LargeBinary, String, UniqueConstraint, func
from sqlalchemy.orm import Mapped, mapped_column

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_SET_NULL
from infrastructure.models.pk import SurrogateKey


class DocumentoXmlORM(Base):
    """Metadados de XML fiscal em disco — sem BLOB."""

    __tablename__ = "documento_xml"
    __table_args__ = (
        UniqueConstraint("chave_nfe", name="uq_documento_xml_chave_nfe"),
        Index("ix_documento_xml_numero_nf", "numero_nf"),
        Index("ix_documento_xml_hash_sha256", "hash_sha256"),
        Index("ix_documento_xml_ativo", "ativo"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    chave_nfe: Mapped[str] = mapped_column(CHAR(44), nullable=False)
    numero_nf: Mapped[str] = mapped_column(String(20), nullable=False)
    nome_arquivo: Mapped[str] = mapped_column(String(255), nullable=False)
    caminho_arquivo: Mapped[str] = mapped_column(String(500), nullable=False)
    hash_sha256: Mapped[str] = mapped_column(String(64), nullable=False)
    tamanho: Mapped[int] = mapped_column(Integer, nullable=False, default=0)
    conteudo_xml: Mapped[bytes | None] = mapped_column(LargeBinary, nullable=True)
    usuario_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    data_importacao: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        index=True,
    )
    ativo: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True, server_default="1")
