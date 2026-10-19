"""Incorporação, cisão e leilão de fração; propostas de evento ignoradas (F7a).

Revision ID: 0005
Revises: 0004
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0005"
down_revision: str | None = "0004"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    with op.batch_alter_table("evento_corporativo", schema=None) as batch_op:
        batch_op.add_column(sa.Column("ativo_destino_id", sa.Integer(), nullable=True))
        batch_op.add_column(sa.Column("quantidade_destino", sa.Text(), nullable=True))
        batch_op.add_column(sa.Column("natureza", sa.String(length=24), nullable=True))
        batch_op.add_column(sa.Column("percentual_custo", sa.Text(), nullable=True))
        batch_op.create_foreign_key(
            batch_op.f("fk_evento_corporativo_ativo_destino_id_ativo"), "ativo",
            ["ativo_destino_id"], ["id"],
        )  # fmt: skip
    op.create_table(
        "proposta_ignorada",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("ignorada_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_proposta_ignorada_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_proposta_ignorada")),
        sa.UniqueConstraint("chave", name=op.f("uq_proposta_ignorada_chave")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("proposta_ignorada")
    with op.batch_alter_table("evento_corporativo", schema=None) as batch_op:
        batch_op.drop_constraint(
            batch_op.f("fk_evento_corporativo_ativo_destino_id_ativo"), type_="foreignkey"
        )
        batch_op.drop_column("percentual_custo")
        batch_op.drop_column("natureza")
        batch_op.drop_column("quantidade_destino")
        batch_op.drop_column("ativo_destino_id")
