from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, Index, String, UniqueConstraint, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_RESTRICT
from infrastructure.models.pk import SurrogateKey


class DocumentoORM(Base):
    """Metadados de PDF em disco — sem BLOB."""

    __tablename__ = "documento"
    __table_args__ = (
        UniqueConstraint("carregamento_id", "tipo", name="uq_documento_carregamento_tipo"),
        Index("ix_documento_tipo_criado_em", "tipo", "criado_em"),
        Index("ix_documento_hash_sha256", "hash_sha256"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    carregamento_id: Mapped[int] = mapped_column(
        ForeignKey("carregamento.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    usuario_id: Mapped[int] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    tipo: Mapped[str] = mapped_column(String(30), nullable=False, index=True)
    caminho_arquivo: Mapped[str] = mapped_column(String(500), nullable=False)
    nome_arquivo: Mapped[str] = mapped_column(String(255), nullable=False)
    hash_sha256: Mapped[str] = mapped_column(String(64), nullable=False)
    criado_em: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        index=True,
    )

    carregamento: Mapped["CarregamentoORM"] = relationship(back_populates="documentos")
