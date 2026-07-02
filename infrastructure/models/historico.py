from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, Index, String, func
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_RESTRICT, ON_DELETE_SET_NULL
from infrastructure.models.pk import SurrogateKey


class HistoricoOperacionalORM(Base):
    """Historico imutavel vinculado a carregamentos (nao permite exclusao fisica)."""

    __tablename__ = "historico_operacional"
    __table_args__ = (
        Index("ix_historico_operacional_evento_criado_em", "evento", "criado_em"),
        Index("ix_historico_operacional_carregamento_criado_em", "carregamento_id", "criado_em"),
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
    item_carregamento_id: Mapped[int | None] = mapped_column(
        ForeignKey("item_carregamento.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    evento: Mapped[str] = mapped_column(String(60), nullable=False, index=True)
    descricao: Mapped[str | None] = mapped_column(String(500), nullable=True)
    criado_em: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        index=True,
    )

    carregamento: Mapped["CarregamentoORM"] = relationship(back_populates="historicos")
