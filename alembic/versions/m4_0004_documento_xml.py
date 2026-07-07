"""Revision M4: tabela documento_xml para metadados de XML fiscal em disco."""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "m4_0004_documento_xml"
down_revision: Union[str, None] = "m3_0003_integer_surrogate_keys"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def _boolean_default_true():
    if op.get_bind().dialect.name == "postgresql":
        return sa.text("true")
    return sa.text("1")


def upgrade() -> None:
    op.create_table(
        "documento_xml",
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("chave_nfe", sa.CHAR(length=44), nullable=False),
        sa.Column("numero_nf", sa.String(length=20), nullable=False),
        sa.Column("nome_arquivo", sa.String(length=255), nullable=False),
        sa.Column("caminho_arquivo", sa.String(length=500), nullable=False),
        sa.Column("hash_sha256", sa.String(length=64), nullable=False),
        sa.Column("tamanho", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("usuario_id", sa.Integer(), nullable=True),
        sa.Column(
            "data_importacao",
            sa.DateTime(timezone=True),
            server_default=sa.text("(CURRENT_TIMESTAMP)"),
            nullable=False,
        ),
        sa.Column("ativo", sa.Boolean(), server_default=_boolean_default_true(), nullable=False),
        sa.ForeignKeyConstraint(["usuario_id"], ["usuario.id"], onupdate="CASCADE", ondelete="SET NULL"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("chave_nfe", name="uq_documento_xml_chave_nfe"),
    )
    op.create_index("ix_documento_xml_numero_nf", "documento_xml", ["numero_nf"], unique=False)
    op.create_index("ix_documento_xml_hash_sha256", "documento_xml", ["hash_sha256"], unique=False)
    op.create_index("ix_documento_xml_ativo", "documento_xml", ["ativo"], unique=False)
    op.create_index("ix_documento_xml_usuario_id", "documento_xml", ["usuario_id"], unique=False)
    op.create_index("ix_documento_xml_data_importacao", "documento_xml", ["data_importacao"], unique=False)


def downgrade() -> None:
    op.drop_index("ix_documento_xml_data_importacao", table_name="documento_xml")
    op.drop_index("ix_documento_xml_usuario_id", table_name="documento_xml")
    op.drop_index("ix_documento_xml_ativo", table_name="documento_xml")
    op.drop_index("ix_documento_xml_hash_sha256", table_name="documento_xml")
    op.drop_index("ix_documento_xml_numero_nf", table_name="documento_xml")
    op.drop_table("documento_xml")
