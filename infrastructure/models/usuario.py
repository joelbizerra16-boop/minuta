from __future__ import annotations

import uuid
from datetime import datetime

from sqlalchemy import Boolean, DateTime, ForeignKey, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.constants import ON_DELETE_RESTRICT, ON_DELETE_SET_NULL
from infrastructure.models.mixins import SoftDeleteMixin, TimestampMixin
from infrastructure.models.pk import SurrogateKey


class UsuarioORM(TimestampMixin, SoftDeleteMixin, Base):
    __tablename__ = "usuario"

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    uuid: Mapped[str] = mapped_column(String(36), nullable=False, unique=True, default=lambda: str(uuid.uuid4()))
    nome: Mapped[str] = mapped_column(String(200), nullable=False)
    # Campo `usuario` equivale ao login do dominio JSON.
    usuario: Mapped[str] = mapped_column(String(80), nullable=False, unique=True)
    senha_hash: Mapped[str] = mapped_column(String(255), nullable=False)
    perfil_id: Mapped[int] = mapped_column(
        ForeignKey("perfil.id", ondelete=ON_DELETE_RESTRICT, onupdate="CASCADE"),
        nullable=False,
        index=True,
    )
    # Snapshot denormalizado (`perfil_snapshot` no dominio).
    perfil: Mapped[str] = mapped_column(String(20), nullable=False, default="OPERADOR")
    bloqueado: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False, server_default="0")
    ultimo_login: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    criado_por_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
    )
    atualizado_por_id: Mapped[int | None] = mapped_column(
        ForeignKey("usuario.id", ondelete=ON_DELETE_SET_NULL, onupdate="CASCADE"),
        nullable=True,
    )

    perfil_ref: Mapped["PerfilORM"] = relationship(back_populates="usuarios")
