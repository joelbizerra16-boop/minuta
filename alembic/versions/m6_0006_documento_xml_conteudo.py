"""Revision M6: conteudo bruto do XML fiscal persistido no banco."""

from __future__ import annotations

from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

revision: str = "m6_0006_documento_xml_conteudo"
down_revision: Union[str, None] = "m5_0005_operational_tables"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "documento_xml",
        sa.Column("conteudo_xml", sa.LargeBinary(), nullable=True),
    )


def downgrade() -> None:
    op.drop_column("documento_xml", "conteudo_xml")
