"""Planilhas da B3 (movimentação, negociação, posição) e outros custos dedutíveis (F7a).

Revision ID: 0004
Revises: 0003
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0004"
down_revision: str | None = "0003"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "arquivo_b3",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("sha256", sa.String(length=64), nullable=False),
        sa.Column("nome", sa.Text(), nullable=False),
        sa.Column("tipo", sa.String(length=20), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("importado_em", sa.DateTime(), nullable=False),
        sa.Column("linhas", sa.Integer(), nullable=False),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_arquivo_b3_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_arquivo_b3")),
        sa.UniqueConstraint("sha256", name=op.f("uq_arquivo_b3_sha256")),
    )
    op.create_table(
        "outro_custo",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("ano", sa.Integer(), nullable=False),
        sa.Column("mes", sa.Integer(), nullable=False),
        sa.Column("categoria", sa.String(length=10), nullable=False),
        sa.Column("valor", sa.Text(), nullable=False),
        sa.Column("descricao", sa.Text(), nullable=False),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_outro_custo_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_outro_custo")),
        sa.UniqueConstraint(
            "titular_id", "ano", "mes", "categoria", name=op.f("uq_outro_custo_titular_id")
        ),
    )
    op.create_table(
        "movimentacao_b3",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("arquivo_id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("data", sa.Date(), nullable=False),
        sa.Column("entrada", sa.Boolean(), nullable=False),
        sa.Column("movimentacao", sa.Text(), nullable=False),
        sa.Column("produto", sa.Text(), nullable=False),
        sa.Column("codigo", sa.String(length=20), nullable=False),
        sa.Column("instituicao", sa.Text(), nullable=False),
        sa.Column("quantidade", sa.Text(), nullable=False),
        sa.Column("preco", sa.Text(), nullable=True),
        sa.Column("valor", sa.Text(), nullable=True),
        sa.ForeignKeyConstraint(
            ["arquivo_id"], ["arquivo_b3.id"], name=op.f("fk_movimentacao_b3_arquivo_id_arquivo_b3")
        ),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_movimentacao_b3_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_movimentacao_b3")),
        sa.UniqueConstraint("chave", name=op.f("uq_movimentacao_b3_chave")),
    )
    with op.batch_alter_table("movimentacao_b3", schema=None) as batch_op:
        batch_op.create_index(
            batch_op.f("ix_movimentacao_b3_titular_id"), ["titular_id"], unique=False
        )

    op.create_table(
        "negociacao_b3",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("chave", sa.String(length=64), nullable=False),
        sa.Column("arquivo_id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("data", sa.Date(), nullable=False),
        sa.Column("cv", sa.String(length=1), nullable=False),
        sa.Column("mercado", sa.Text(), nullable=False),
        sa.Column("vencimento", sa.Date(), nullable=True),
        sa.Column("instituicao", sa.Text(), nullable=False),
        sa.Column("codigo", sa.String(length=20), nullable=False),
        sa.Column("quantidade", sa.Text(), nullable=False),
        sa.Column("preco", sa.Text(), nullable=False),
        sa.Column("valor", sa.Text(), nullable=False),
        sa.ForeignKeyConstraint(
            ["arquivo_id"], ["arquivo_b3.id"], name=op.f("fk_negociacao_b3_arquivo_id_arquivo_b3")
        ),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_negociacao_b3_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_negociacao_b3")),
        sa.UniqueConstraint("chave", name=op.f("uq_negociacao_b3_chave")),
    )
    with op.batch_alter_table("negociacao_b3", schema=None) as batch_op:
        batch_op.create_index(
            batch_op.f("ix_negociacao_b3_titular_id"), ["titular_id"], unique=False
        )

    op.create_table(
        "posicao_b3",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("arquivo_id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("data", sa.Date(), nullable=False),
        sa.Column("instituicao", sa.Text(), nullable=False),
        sa.Column("conta", sa.String(length=40), nullable=False),
        sa.Column("codigo", sa.String(length=20), nullable=False),
        sa.Column("isin", sa.String(length=12), nullable=False),
        sa.Column("quantidade", sa.Text(), nullable=False),
        sa.Column("origem", sa.String(length=12), nullable=False),
        sa.ForeignKeyConstraint(
            ["arquivo_id"], ["arquivo_b3.id"], name=op.f("fk_posicao_b3_arquivo_id_arquivo_b3")
        ),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_posicao_b3_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_posicao_b3")),
    )
    with op.batch_alter_table("posicao_b3", schema=None) as batch_op:
        batch_op.create_index(batch_op.f("ix_posicao_b3_titular_id"), ["titular_id"], unique=False)


def downgrade() -> None:
    with op.batch_alter_table("posicao_b3", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_posicao_b3_titular_id"))

    op.drop_table("posicao_b3")
    with op.batch_alter_table("negociacao_b3", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_negociacao_b3_titular_id"))

    op.drop_table("negociacao_b3")
    with op.batch_alter_table("movimentacao_b3", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_movimentacao_b3_titular_id"))

    op.drop_table("movimentacao_b3")
    op.drop_table("outro_custo")
    op.drop_table("arquivo_b3")
