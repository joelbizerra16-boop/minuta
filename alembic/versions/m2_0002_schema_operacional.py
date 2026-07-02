"""Revision M2: schema operacional completo (tabelas M0.5 fora de M1)."""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "m2_0002_schema_operacional"
down_revision: Union[str, None] = "m1_0001_perfil_usuario"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "configuracao",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("chave", sa.String(length=120), nullable=False),
        sa.Column("valor", sa.Text(), nullable=False, server_default=""),
        sa.Column("categoria", sa.String(length=60), nullable=False, server_default="GERAL"),
        sa.Column("tipo_valor", sa.String(length=20), nullable=False, server_default="STRING"),
        sa.Column("descricao", sa.String(length=255), nullable=True),
        sa.Column("atualizado_por_usuario_id", sa.Integer(), nullable=True),
        sa.Column("criado_em", sa.DateTime(timezone=True), server_default=sa.text("(CURRENT_TIMESTAMP)"), nullable=False),
        sa.Column("atualizado_em", sa.DateTime(timezone=True), server_default=sa.text("(CURRENT_TIMESTAMP)"), nullable=False),
        sa.ForeignKeyConstraint(
            ["atualizado_por_usuario_id"],
            ["usuario.id"],
            onupdate="CASCADE",
            ondelete="SET NULL",
        ),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("chave", name="uq_configuracao_chave"),
    )
    op.create_index("ix_configuracao_categoria", "configuracao", ["categoria"], unique=False)
    op.create_index("ix_configuracao_atualizado_por_usuario_id", "configuracao", ["atualizado_por_usuario_id"], unique=False)

    # Demais tabelas M0.5: em ambientes novos use `ensure_full_schema()` ou autogenerate
    # complementar. Esta revision formaliza a tabela que substitui os JSONs operacionais.


def downgrade() -> None:
    op.drop_index("ix_configuracao_atualizado_por_usuario_id", table_name="configuracao")
    op.drop_index("ix_configuracao_categoria", table_name="configuracao")
    op.drop_table("configuracao")
