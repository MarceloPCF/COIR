"""Imposto abaixo do mínimo da DARF acumulado em 31/12 no saldo inicial (Lei 9.430/1996 art. 68
§1º; decisão do usuário de 10/11/2026).

Revision ID: 0010
Revises: 0009
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0010"
down_revision: str | None = "0009"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    with op.batch_alter_table("prejuizo_inicial") as tabela:
        tabela.add_column(
            sa.Column("darf_acumulado", sa.Text(), nullable=False, server_default="0")
        )


def downgrade() -> None:
    with op.batch_alter_table("prejuizo_inicial") as tabela:
        tabela.drop_column("darf_acumulado")
