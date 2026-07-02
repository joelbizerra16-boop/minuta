from __future__ import annotations

from datetime import datetime

from sqlalchemy import Boolean, DateTime, func
from sqlalchemy.orm import Mapped, mapped_column


class TimestampMixin:
    criado_em: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
    )
    atualizado_em: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )


class SoftDeleteMixin:
    """Exclusao logica padrao para cadastros permanentes."""

    ativo: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True, server_default="1")
    excluido_em: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
