"""Revision M3: chaves substitutas Integer (autoincrement nativo SQLite/PostgreSQL)."""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "m3_0003_integer_surrogate_keys"
down_revision: Union[str, None] = "m2_0002_schema_operacional"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

# Colunas BIGINT legadas a converter em INTEGER no PostgreSQL.
_PG_INTEGER_COLUMNS: list[tuple[str, str]] = [
    ("perfil", "id"),
    ("usuario", "id"),
    ("usuario", "perfil_id"),
    ("usuario", "criado_por_id"),
    ("usuario", "atualizado_por_id"),
    ("configuracao", "id"),
    ("configuracao", "atualizado_por_usuario_id"),
    ("motorista", "id"),
    ("veiculo", "id"),
    ("destinatario", "id"),
    ("rota", "id"),
    ("nota_fiscal", "id"),
    ("nota_fiscal", "destinatario_id"),
    ("nota_fiscal", "rota_id"),
    ("item_nota_fiscal", "id"),
    ("item_nota_fiscal", "nota_fiscal_id"),
    ("carregamento", "id"),
    ("carregamento", "usuario_id"),
    ("carregamento", "motorista_id"),
    ("carregamento", "veiculo_id"),
    ("carregamento", "ultima_impressao_usuario_id"),
    ("item_carregamento", "id"),
    ("item_carregamento", "carregamento_id"),
    ("item_carregamento", "nota_fiscal_id"),
    ("documento", "id"),
    ("documento", "carregamento_id"),
    ("documento", "usuario_id"),
    ("historico_operacional", "id"),
    ("historico_operacional", "carregamento_id"),
    ("historico_operacional", "usuario_id"),
    ("historico_operacional", "item_carregamento_id"),
    ("evento_auditoria", "id"),
    ("evento_auditoria", "usuario_id"),
    ("evento_auditoria", "entidade_id"),
]


def _is_postgresql() -> bool:
    bind = op.get_bind()
    return bind.dialect.name == "postgresql"


def upgrade() -> None:
    if not _is_postgresql():
        # SQLite: novos ambientes usam create_all/ensure_full_schema com Integer.
        # Bancos SQLite legados com BIGINT devem ser recriados em desenvolvimento.
        return

    for table_name, column_name in _PG_INTEGER_COLUMNS:
        op.alter_column(
            table_name,
            column_name,
            existing_type=sa.BigInteger(),
            type_=sa.Integer(),
            postgresql_using=f"{column_name}::integer",
        )


def downgrade() -> None:
    if not _is_postgresql():
        return

    for table_name, column_name in reversed(_PG_INTEGER_COLUMNS):
        op.alter_column(
            table_name,
            column_name,
            existing_type=sa.Integer(),
            type_=sa.BigInteger(),
            postgresql_using=f"{column_name}::bigint",
        )
