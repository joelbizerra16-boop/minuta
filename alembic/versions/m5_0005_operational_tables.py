"""Revision M5: tabelas operacionais M0.5 (cadastros, NF, carregamento, auditoria)."""

from __future__ import annotations

from typing import Sequence, Union

from alembic import op
from sqlalchemy import inspect

revision: str = "m5_0005_operational_tables"
down_revision: Union[str, None] = "m4_0004_documento_xml"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

# Ordem respeitando dependencias de FK.
_OPERATIONAL_TABLES: tuple[str, ...] = (
    "motorista",
    "veiculo",
    "destinatario",
    "rota",
    "nota_fiscal",
    "item_nota_fiscal",
    "carregamento",
    "item_carregamento",
    "documento",
    "historico_operacional",
    "evento_auditoria",
)


def upgrade() -> None:
    bind = op.get_bind()
    inspector = inspect(bind)
    existing = set(inspector.get_table_names())

    # Garante registro completo do metadata ORM.
    from infrastructure.models import (  # noqa: F401
        Base,
        CarregamentoORM,
        DestinatarioORM,
        DocumentoORM,
        EventoAuditoriaORM,
        HistoricoOperacionalORM,
        ItemCarregamentoORM,
        ItemNotaFiscalORM,
        MotoristaORM,
        NotaFiscalORM,
        RotaORM,
        VeiculoORM,
    )

    for table_name in _OPERATIONAL_TABLES:
        if table_name in existing:
            continue
        Base.metadata.tables[table_name].create(bind, checkfirst=True)


def downgrade() -> None:
    bind = op.get_bind()
    for table_name in reversed(_OPERATIONAL_TABLES):
        if table_name in inspect(bind).get_table_names():
            op.drop_table(table_name)
