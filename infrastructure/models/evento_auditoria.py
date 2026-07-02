from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, Index, String, Text, func
from sqlalchemy.orm import Mapped, mapped_column

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_SET_NULL
from infrastructure.models.pk import SurrogateKey


class EventoAuditoriaORM(Base):
    """
    Trilha de auditoria transversal do sistema.
    Modelada na M0.5; utilizacao prevista para fases posteriores da migracao.
    """

    __tablename__ = "evento_auditoria"
    __table_args__ = (
        Index("ix_evento_auditoria_categoria_evento", "categoria", "evento"),
        Index("ix_evento_auditoria_entidade", "entidade_tipo", "entidade_id"),
        Index("ix_evento_auditoria_usuario_criado_em", "usuario_id", "criado_em"),
        Index("ix_evento_auditoria_criado_em", "criado_em"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    usuario_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
    categoria: Mapped[str] = mapped_column(String(40), nullable=False, index=True)
    evento: Mapped[str] = mapped_column(String(60), nullable=False, index=True)
    entidade_tipo: Mapped[str | None] = mapped_column(String(40), nullable=True)
    entidade_id: Mapped[int | None] = mapped_column(SurrogateKey, nullable=True)
    descricao: Mapped[str | None] = mapped_column(String(500), nullable=True)
    metadados_json: Mapped[str | None] = mapped_column(Text, nullable=True)
    ip_origem: Mapped[str | None] = mapped_column(String(45), nullable=True)
    criado_em: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
    )
