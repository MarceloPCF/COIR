"""Itens da declaração anual marcados como lançados no programa IRPF (F8, tela 14).

Revision ID: 0007
Revises: 0006
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0007"
down_revision: str | None = "0006"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "irpf_lancado",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("ano", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=120), nullable=False),
        sa.Column("marcado_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_irpf_lancado_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_irpf_lancado")),
        sa.UniqueConstraint("titular_id", "ano", "chave", name=op.f("uq_irpf_lancado_titular_id")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("irpf_lancado")
