"""Versão anterior da nota que o reprocessamento trocou ou removeu.

Revision ID: 0009
Revises: 0008
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0009"
down_revision: str | None = "0008"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "nota_anterior",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("arquivo_id", sa.Integer(), nullable=False),
        sa.Column("conta_id", sa.Integer(), nullable=False),
        sa.Column("numero", sa.String(length=40), nullable=False),
        sa.Column("data_pregao", sa.Date(), nullable=False),
        sa.Column("segmento", sa.String(length=10), nullable=False),
        sa.Column("paginas", sa.Text(), nullable=False),
        sa.Column("json_nota", sa.Text(), nullable=False),
        sa.Column("motivo", sa.String(length=12), nullable=False),
        sa.Column("mudou", sa.Text(), nullable=False),
        sa.Column("alterada_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["arquivo_id"], ["arquivo_importado.id"],
                                name=op.f("fk_nota_anterior_arquivo_id_arquivo_importado")),
        sa.ForeignKeyConstraint(["conta_id"], ["conta.id"],
                                name=op.f("fk_nota_anterior_conta_id_conta")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_nota_anterior")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("nota_anterior")
