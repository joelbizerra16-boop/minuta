from __future__ import annotations

from sqlalchemy import String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from infrastructure.models.base import Base
from infrastructure.models.pk import SurrogateKey


class PerfilORM(Base):
    __tablename__ = "perfil"

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    codigo: Mapped[str] = mapped_column(String(20), nullable=False, unique=True)
    nome: Mapped[str] = mapped_column(String(80), nullable=False)

    usuarios: Mapped[list["UsuarioORM"]] = relationship(back_populates="perfil_ref")
