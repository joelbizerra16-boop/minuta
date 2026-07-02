"""Revision M1: perfil e usuario."""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "m1_0001_perfil_usuario"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "perfil",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("codigo", sa.String(length=20), nullable=False),
        sa.Column("nome", sa.String(length=80), nullable=False),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("codigo", name="uq_perfil_codigo"),
    )

    perfil_table = sa.table(
        "perfil",
        sa.column("id", sa.Integer),
        sa.column("codigo", sa.String),
        sa.column("nome", sa.String),
    )
    op.bulk_insert(
        perfil_table,
        [
            {"id": 1, "codigo": "ADMIN", "nome": "Administrador"},
            {"id": 2, "codigo": "OPERADOR", "nome": "Operador"},
        ],
    )

    op.create_table(
        "usuario",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("uuid", sa.String(length=36), nullable=False),
        sa.Column("nome", sa.String(length=200), nullable=False),
        sa.Column("usuario", sa.String(length=80), nullable=False),
        sa.Column("senha_hash", sa.String(length=255), nullable=False),
        sa.Column("perfil_id", sa.Integer(), nullable=False),
        sa.Column("perfil", sa.String(length=20), nullable=False, server_default="OPERADOR"),
        sa.Column("bloqueado", sa.Boolean(), nullable=False, server_default=sa.text("0")),
        sa.Column("ativo", sa.Boolean(), nullable=False, server_default=sa.text("1")),
        sa.Column("excluido_em", sa.DateTime(timezone=True), nullable=True),
        sa.Column("criado_em", sa.DateTime(timezone=True), server_default=sa.text("(CURRENT_TIMESTAMP)"), nullable=False),
        sa.Column("atualizado_em", sa.DateTime(timezone=True), server_default=sa.text("(CURRENT_TIMESTAMP)"), nullable=False),
        sa.Column("ultimo_login", sa.DateTime(timezone=True), nullable=True),
        sa.Column("criado_por_id", sa.Integer(), nullable=True),
        sa.Column("atualizado_por_id", sa.Integer(), nullable=True),
        sa.ForeignKeyConstraint(["atualizado_por_id"], ["usuario.id"], onupdate="CASCADE", ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["criado_por_id"], ["usuario.id"], onupdate="CASCADE", ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["perfil_id"], ["perfil.id"], onupdate="CASCADE", ondelete="RESTRICT"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("usuario", name="uq_usuario_login"),
        sa.UniqueConstraint("uuid", name="uq_usuario_uuid"),
    )
    op.create_index("ix_usuario_perfil_id", "usuario", ["perfil_id"], unique=False)


def downgrade() -> None:
    op.drop_index("ix_usuario_perfil_id", table_name="usuario")
    op.drop_table("usuario")
    op.drop_table("perfil")
