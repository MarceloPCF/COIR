"""Valor de mercado na entrega do dividendo pago em bens (parecer de 16/11/2026).

Revision ID: 0012
Revises: 0011
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0012"
down_revision: str | None = "0011"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    with op.batch_alter_table("evento_corporativo") as tabela:
        tabela.add_column(sa.Column("valor_corrente", sa.Text(), nullable=True))


def downgrade() -> None:
    with op.batch_alter_table("evento_corporativo") as tabela:
        tabela.drop_column("valor_corrente")
