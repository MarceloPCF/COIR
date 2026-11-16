"""Código escolhido pelo usuário para um nome de pregão que valia para mais de um na data.

Revision ID: 0011
Revises: 0010
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0011"
down_revision: str | None = "0010"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "codigo_escolhido",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("corretora", sa.String(length=40), nullable=False),
        sa.Column("conta", sa.String(length=40), nullable=False),
        sa.Column("nome", sa.Text(), nullable=False),
        sa.Column("ticker", sa.String(length=20), nullable=False),
        sa.Column("registrado_em", sa.DateTime(), nullable=False),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_codigo_escolhido")),
        sa.UniqueConstraint("corretora", "conta", "nome",
                            name=op.f("uq_codigo_escolhido_corretora")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("codigo_escolhido")
