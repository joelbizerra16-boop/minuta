from __future__ import annotations

from sqlalchemy import Index, String, text
from sqlalchemy.orm import Mapped, mapped_column

from infrastructure.models.base import Base
from infrastructure.models.mixins import SoftDeleteMixin, TimestampMixin
from infrastructure.models.pk import SurrogateKey


class MotoristaORM(TimestampMixin, SoftDeleteMixin, Base):
    __tablename__ = "motorista"
    __table_args__ = (
        Index(
            "uq_motorista_nome_ativo",
            "nome",
            unique=True,
            sqlite_where=text("ativo = 1 AND excluido_em IS NULL"),
            postgresql_where=text("ativo = true AND excluido_em IS NULL"),
        ),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    nome: Mapped[str] = mapped_column(String(200), nullable=False, index=True)


class VeiculoORM(TimestampMixin, SoftDeleteMixin, Base):
    __tablename__ = "veiculo"
    __table_args__ = (
        Index(
            "uq_veiculo_placa_ativo",
            "placa",
            unique=True,
            sqlite_where=text("ativo = 1 AND excluido_em IS NULL"),
            postgresql_where=text("ativo = true AND excluido_em IS NULL"),
        ),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    placa: Mapped[str] = mapped_column(String(20), nullable=False, index=True)


class DestinatarioORM(TimestampMixin, SoftDeleteMixin, Base):
    __tablename__ = "destinatario"
    __table_args__ = (
        Index(
            "uq_destinatario_razao_municipio_uf_ativo",
            "razao_social",
            "municipio",
            "uf",
            unique=True,
            sqlite_where=text("ativo = 1 AND excluido_em IS NULL"),
            postgresql_where=text("ativo = true AND excluido_em IS NULL"),
        ),
        Index("ix_destinatario_municipio_uf", "municipio", "uf"),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    razao_social: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    municipio: Mapped[str | None] = mapped_column(String(120), nullable=True)
    uf: Mapped[str | None] = mapped_column(String(2), nullable=True)


class RotaORM(TimestampMixin, SoftDeleteMixin, Base):
    __tablename__ = "rota"
    __table_args__ = (
        Index(
            "uq_rota_nome_ativo",
            "nome",
            unique=True,
            sqlite_where=text("ativo = 1 AND excluido_em IS NULL"),
            postgresql_where=text("ativo = true AND excluido_em IS NULL"),
        ),
    )

    id: Mapped[int] = mapped_column(SurrogateKey, primary_key=True, autoincrement=True)
    nome: Mapped[str] = mapped_column(String(120), nullable=False, index=True)
