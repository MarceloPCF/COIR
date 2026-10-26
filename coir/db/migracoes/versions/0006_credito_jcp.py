"""Data do crédito de JCP informada pelo usuário (F8).

Revision ID: 0006
Revises: 0005
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0006"
down_revision: str | None = "0005"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "credito_jcp",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("data_credito", sa.Date(), nullable=False),
        sa.Column("registrado_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_credito_jcp_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_credito_jcp")),
        sa.UniqueConstraint("chave", name=op.f("uq_credito_jcp_chave")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("credito_jcp")
