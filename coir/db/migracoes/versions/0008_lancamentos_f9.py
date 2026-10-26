"""Lançamentos da F9: origem de um reembolso de BTC (dividendo ou JCP) informada pelo usuário e o
prejuízo a compensar em 31/12 de quem começa sem o histórico de notas.

Revision ID: 0008
Revises: 0007
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0008"
down_revision: str | None = "0007"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "origem_reembolso",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("origem", sa.String(length=10), nullable=False),
        sa.Column("registrado_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_origem_reembolso_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_origem_reembolso")),
        sa.UniqueConstraint("chave", name=op.f("uq_origem_reembolso_chave")),
    )  # fmt: skip

    op.create_table(
        "prejuizo_inicial",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("ano", sa.Integer(), nullable=False),
        sa.Column("comum", sa.Text(), nullable=False),
        sa.Column("day_trade", sa.Text(), nullable=False),
        sa.Column("fii", sa.Text(), nullable=False),
        sa.Column("registrado_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_prejuizo_inicial_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_prejuizo_inicial")),
        sa.UniqueConstraint("titular_id", name=op.f("uq_prejuizo_inicial_titular_id")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("prejuizo_inicial")
    op.drop_table("origem_reembolso")
