from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, ForeignKey, String, Text, func
from sqlalchemy.orm import Mapped, mapped_column

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_SET_NULL
from infrastructure.models.mixins import TimestampMixin
from infrastructure.models.pk import SurrogateKey


class ConfiguracaoORM(TimestampMixin, Base):
    """Parametros permanentes da aplicacao (substituicao futura de JSONs de configuracao)."""

    __tablename__ = "configuracao"

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    chave: Mapped[str] = mapped_column(String(120), nullable=False, unique=True)
    valor: Mapped[str] = mapped_column(Text, nullable=False, default="")
    categoria: Mapped[str] = mapped_column(String(60), nullable=False, default="GERAL", index=True)
    tipo_valor: Mapped[str] = mapped_column(String(20), nullable=False, default="STRING")
    descricao: Mapped[str | None] = mapped_column(String(255), nullable=True)
    atualizado_por_usuario_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
        index=True,
    )
