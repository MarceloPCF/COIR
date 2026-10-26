"""Eventos corporativos e de custódia (F4).

Revision ID: 0002
Revises: 0001
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0002"
down_revision: str | None = "0001"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "evento_corporativo",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("data", sa.Date(), nullable=False),
        sa.Column("conta_id", sa.Integer(), nullable=False),
        sa.Column("ativo_id", sa.Integer(), nullable=False),
        sa.Column("tipo", sa.String(length=20), nullable=False),
        sa.Column("cv", sa.String(length=1), nullable=True),
        sa.Column("quantidade", sa.Text(), nullable=True),
        sa.Column("valor", sa.Text(), nullable=False),
        sa.Column("conta_destino_id", sa.Integer(), nullable=True),
        sa.Column("observacao", sa.Text(), nullable=False),
        sa.Column("origem", sa.String(length=20), nullable=False),
        sa.ForeignKeyConstraint(["ativo_id"], ["ativo.id"],
                                name=op.f("fk_evento_corporativo_ativo_id_ativo")),
        sa.ForeignKeyConstraint(["conta_destino_id"], ["conta.id"],
                                name=op.f("fk_evento_corporativo_conta_destino_id_conta")),
        sa.ForeignKeyConstraint(["conta_id"], ["conta.id"],
                                name=op.f("fk_evento_corporativo_conta_id_conta")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_evento_corporativo")),
        sa.UniqueConstraint("chave", name=op.f("uq_evento_corporativo_chave")),
    )  # fmt: skip
    with op.batch_alter_table("evento_corporativo", schema=None) as batch_op:
        batch_op.create_index(batch_op.f("ix_evento_corporativo_ativo_id"), ["ativo_id"])
        batch_op.create_index(batch_op.f("ix_evento_corporativo_conta_id"), ["conta_id"])


def downgrade() -> None:
    with op.batch_alter_table("evento_corporativo", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_evento_corporativo_conta_id"))
        batch_op.drop_index(batch_op.f("ix_evento_corporativo_ativo_id"))
    op.drop_table("evento_corporativo")
