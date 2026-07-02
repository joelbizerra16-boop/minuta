from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

from sqlalchemy import (
    CHAR,
    CheckConstraint,
    Date,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    Numeric,
    String,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_RESTRICT, ON_DELETE_SET_NULL
from infrastructure.models.mixins import TimestampMixin
from infrastructure.models.pk import SurrogateKey


class NotaFiscalORM(TimestampMixin, Base):
    __tablename__ = "nota_fiscal"
    __table_args__ = (
        UniqueConstraint("chave_nfe", name="uq_nota_fiscal_chave_nfe"),
        Index("ix_nota_fiscal_numero_nf_emitente", "numero_nf", "emitente"),
        Index("ix_nota_fiscal_destinatario_status", "destinatario", "status_nf"),
        Index("ix_nota_fiscal_rota_data_emissao", "rota", "data_emissao"),
        CheckConstraint("length(chave_nfe) = 44", name="ck_nota_fiscal_chave_nfe_len"),
        CheckConstraint("valor_total >= 0", name="ck_nota_fiscal_valor_total_nonneg"),
        CheckConstraint("peso_total >= 0", name="ck_nota_fiscal_peso_total_nonneg"),
        CheckConstraint("volume_total >= 0", name="ck_nota_fiscal_volume_total_nonneg"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    chave_nfe: Mapped[str] = mapped_column(CHAR(44), nullable=False)
    numero_nf: Mapped[str] = mapped_column(String(20), nullable=False, index=True)
    emitente: Mapped[str | None] = mapped_column(String(255), nullable=True)
    destinatario_id: Mapped[int | None] = mapped_column(
        ForeignKey("destinatario.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    rota_id: Mapped[int | None] = mapped_column(
        ForeignKey("rota.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    destinatario: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    municipio: Mapped[str | None] = mapped_column(String(120), nullable=True)
    uf: Mapped[str | None] = mapped_column(String(2), nullable=True)
    rota: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    status_nf: Mapped[str] = mapped_column(String(120), nullable=False, index=True)
    tipo_xml: Mapped[str | None] = mapped_column(String(40), nullable=True)
    data_emissao: Mapped[date | None] = mapped_column(Date, nullable=True, index=True)
    data_referencia: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    valor_total: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)
    peso_total: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)
    volume_total: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)
    arquivo_origem: Mapped[str | None] = mapped_column(String(255), nullable=True)
    # criado_em / atualizado_em herdados de TimestampMixin (importacao XML).
    versao: Mapped[int] = mapped_column(Integer, nullable=False, default=1, server_default="1")

    itens: Mapped[list["ItemNotaFiscalORM"]] = relationship(
        back_populates="nota_fiscal",
        cascade="save-update, merge",
        passive_deletes=True,
    )


class ItemNotaFiscalORM(Base):
    __tablename__ = "item_nota_fiscal"
    __table_args__ = (
        UniqueConstraint("nota_fiscal_id", "sequencia", name="uq_item_nota_fiscal_nf_sequencia"),
        Index("ix_item_nota_fiscal_codigo_produto", "codigo_produto"),
        Index("ix_item_nota_fiscal_nf_codigo", "nota_fiscal_id", "codigo_produto"),
        CheckConstraint("quantidade >= 0", name="ck_item_nota_fiscal_quantidade_nonneg"),
        CheckConstraint("peso >= 0", name="ck_item_nota_fiscal_peso_nonneg"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    nota_fiscal_id: Mapped[int] = mapped_column(
        ForeignKey("nota_fiscal.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    sequencia: Mapped[int] = mapped_column(nullable=False, default=1)
    codigo_produto: Mapped[str] = mapped_column(String(60), nullable=False)
    descricao: Mapped[str] = mapped_column(String(500), nullable=False)
    quantidade: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)
    unidade: Mapped[str | None] = mapped_column(String(20), nullable=True)
    peso: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)

    nota_fiscal: Mapped[NotaFiscalORM] = relationship(back_populates="itens")
