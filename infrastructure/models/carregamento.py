from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal

from sqlalchemy import (
    Boolean,
    CheckConstraint,
    Date,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    Numeric,
    String,
    Time,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_RESTRICT, ON_DELETE_SET_NULL
from infrastructure.models.mixins import TimestampMixin
from infrastructure.models.pk import SurrogateKey


class CarregamentoORM(TimestampMixin, Base):
    __tablename__ = "carregamento"
    __table_args__ = (
        UniqueConstraint("numero_carregamento", name="uq_carregamento_numero"),
        Index("ix_carregamento_data_status", "data", "status"),
        Index("ix_carregamento_motorista_data", "motorista", "data"),
        Index("ix_carregamento_placa_data", "placa", "data"),
        Index("ix_carregamento_usuario_data", "usuario_id", "data"),
        Index("ix_carregamento_modalidade_status", "modalidade", "status"),
        CheckConstraint("quantidade_nf >= 0", name="ck_carregamento_quantidade_nf_nonneg"),
        CheckConstraint("quantidade_itens >= 0", name="ck_carregamento_quantidade_itens_nonneg"),
        CheckConstraint("peso_total >= 0", name="ck_carregamento_peso_total_nonneg"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    numero_carregamento: Mapped[str] = mapped_column(String(80), nullable=False)
    usuario_id: Mapped[int] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    motorista_id: Mapped[int | None] = mapped_column(
        ForeignKey("motorista.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    veiculo_id: Mapped[int | None] = mapped_column(
        ForeignKey("veiculo.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    data: Mapped[date] = mapped_column(Date, nullable=False, index=True)
    hora: Mapped[time] = mapped_column(Time, nullable=False)
    # Snapshots denormalizados para historico imutavel mesmo se cadastro for alterado.
    motorista: Mapped[str | None] = mapped_column(String(200), nullable=True, index=True)
    placa: Mapped[str | None] = mapped_column(String(20), nullable=True, index=True)
    filial: Mapped[str | None] = mapped_column(String(120), nullable=True)
    data_saida: Mapped[str | None] = mapped_column(String(30), nullable=True)
    modalidade: Mapped[str] = mapped_column(String(20), nullable=False, index=True)
    status: Mapped[str] = mapped_column(String(20), nullable=False, index=True)
    reentrega: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    quantidade_nf: Mapped[int] = mapped_column(nullable=False, default=0)
    quantidade_itens: Mapped[int] = mapped_column(nullable=False, default=0)
    peso_total: Mapped[Decimal] = mapped_column(Numeric(18, 4), nullable=False, default=0)
    versao: Mapped[int] = mapped_column(Integer, nullable=False, default=1, server_default="1")
    quantidade_impressoes: Mapped[int] = mapped_column(Integer, nullable=False, default=0, server_default="0")
    ultima_impressao_em: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    ultima_impressao_usuario_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )

    itens: Mapped[list["ItemCarregamentoORM"]] = relationship(
        back_populates="carregamento",
        cascade="save-update, merge",
        passive_deletes=True,
    )
    documentos: Mapped[list["DocumentoORM"]] = relationship(
        back_populates="carregamento",
        cascade="save-update, merge",
        passive_deletes=True,
    )
    historicos: Mapped[list["HistoricoOperacionalORM"]] = relationship(
        back_populates="carregamento",
        cascade="save-update, merge",
        passive_deletes=True,
    )


class ItemCarregamentoORM(Base):
    __tablename__ = "item_carregamento"
    __table_args__ = (
        UniqueConstraint(
            "carregamento_id",
            "numero_nf",
            "codigo_produto",
            "sequencia",
            name="uq_item_carregamento_linha",
        ),
        Index("ix_item_carregamento_chave_nfe", "chave_nfe"),
        Index("ix_item_carregamento_numero_nf", "numero_nf"),
        Index("ix_item_carregamento_codigo_produto", "codigo_produto"),
        Index("ix_item_carregamento_destinatario_rota", "destinatario", "rota"),
        Index("ix_item_carregamento_nf_chave", "numero_nf", "chave_nfe"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    carregamento_id: Mapped[int] = mapped_column(
        ForeignKey("carregamento.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    nota_fiscal_id: Mapped[int | None] = mapped_column(
        ForeignKey("nota_fiscal.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    chave_nfe: Mapped[str | None] = mapped_column(String(44), nullable=True)
    numero_nf: Mapped[str] = mapped_column(String(20), nullable=False)
    codigo_produto: Mapped[str | None] = mapped_column(String(60), nullable=True)
    descricao: Mapped[str | None] = mapped_column(String(500), nullable=True)
    quantidade: Mapped[Decimal | None] = mapped_column(Numeric(18, 4), nullable=True)
    unidade: Mapped[str | None] = mapped_column(String(20), nullable=True)
    peso: Mapped[Decimal | None] = mapped_column(Numeric(18, 4), nullable=True)
    destinatario: Mapped[str | None] = mapped_column(String(255), nullable=True, index=True)
    rota: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    status_nf: Mapped[str | None] = mapped_column(String(120), nullable=True)
    sequencia: Mapped[int] = mapped_column(nullable=False, default=1, server_default="1")

    carregamento: Mapped[CarregamentoORM] = relationship(back_populates="itens")
