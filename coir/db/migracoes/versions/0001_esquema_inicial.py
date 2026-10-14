"""Esquema inicial (F3): cadastros, arquivos importados, notas, custos e operações.

Decimal como TEXT (coir.db.tipos.DecimalTexto no modelo).

Revision ID: 0001
Revises:
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0001"
down_revision: str | None = None
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "arquivo_importado",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("sha256", sa.String(length=64), nullable=False),
        sa.Column("nome", sa.Text(), nullable=False),
        sa.Column("tamanho", sa.Integer(), nullable=False),
        sa.Column("importado_em", sa.DateTime(), nullable=False),
        sa.Column("completo", sa.Boolean(), nullable=False),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_arquivo_importado")),
        sa.UniqueConstraint("sha256", name=op.f("uq_arquivo_importado_sha256")),
    )
    op.create_table(
        "ativo",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("isin", sa.String(length=12), nullable=True),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_ativo")),
        sa.UniqueConstraint("isin", name=op.f("uq_ativo_isin")),
    )
    op.create_table(
        "corretora",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("nome", sa.String(length=40), nullable=False),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_corretora")),
        sa.UniqueConstraint("nome", name=op.f("uq_corretora_nome")),
    )
    op.create_table(
        "serie_opcao",
        sa.Column("serie", sa.String(length=20), nullable=False),
        sa.Column("classe", sa.String(length=5), nullable=False),
        sa.PrimaryKeyConstraint("serie", name=op.f("pk_serie_opcao")),
    )
    op.create_table(
        "titular",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("cpf", sa.String(length=11), nullable=False),
        sa.Column("nome", sa.Text(), nullable=False),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_titular")),
        sa.UniqueConstraint("cpf", name=op.f("uq_titular_cpf")),
    )
    op.create_table(
        "conta",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("corretora_id", sa.Integer(), nullable=False),
        sa.Column("numero", sa.String(length=40), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.ForeignKeyConstraint(
            ["corretora_id"], ["corretora.id"], name=op.f("fk_conta_corretora_id_corretora")
        ),
        sa.ForeignKeyConstraint(
            ["titular_id"], ["titular.id"], name=op.f("fk_conta_titular_id_titular")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_conta")),
        sa.UniqueConstraint("corretora_id", "numero", name=op.f("uq_conta_corretora_id")),
    )
    op.create_table(
        "ticker_historico",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("ativo_id", sa.Integer(), nullable=False),
        sa.Column("ticker", sa.String(length=20), nullable=False),
        sa.Column("inicio", sa.Date(), nullable=True),
        sa.Column("fim", sa.Date(), nullable=True),
        sa.ForeignKeyConstraint(
            ["ativo_id"], ["ativo.id"], name=op.f("fk_ticker_historico_ativo_id_ativo")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_ticker_historico")),
        sa.UniqueConstraint("ticker", "inicio", name=op.f("uq_ticker_historico_ticker")),
    )
    with op.batch_alter_table("ticker_historico", schema=None) as batch_op:
        batch_op.create_index("ix_ticker_historico_ticker", ["ticker"], unique=False)

    op.create_table(
        "nota_corretagem",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("arquivo_id", sa.Integer(), nullable=False),
        sa.Column("conta_id", sa.Integer(), nullable=False),
        sa.Column("numero", sa.String(length=40), nullable=False),
        sa.Column("data_pregao", sa.Date(), nullable=False),
        sa.Column("segmento", sa.String(length=10), nullable=False),
        sa.Column("paginas", sa.Text(), nullable=False),
        sa.Column("hash_conteudo", sa.String(length=64), nullable=False),
        sa.Column("json_nota", sa.Text(), nullable=False),
        sa.Column("valor_operacoes", sa.Text(), nullable=False),
        sa.Column("custos_total", sa.Text(), nullable=False),
        sa.Column("irrf", sa.Text(), nullable=False),
        sa.Column("irrf_day_trade", sa.Text(), nullable=False),
        sa.Column("liquido", sa.Text(), nullable=False),
        sa.Column("data_liquidacao", sa.Date(), nullable=True),
        sa.ForeignKeyConstraint(
            ["arquivo_id"],
            ["arquivo_importado.id"],
            name=op.f("fk_nota_corretagem_arquivo_id_arquivo_importado"),
        ),
        sa.ForeignKeyConstraint(
            ["conta_id"], ["conta.id"], name=op.f("fk_nota_corretagem_conta_id_conta")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_nota_corretagem")),
        sa.UniqueConstraint(
            "conta_id",
            "numero",
            "data_pregao",
            "segmento",
            name=op.f("uq_nota_corretagem_conta_id"),
        ),
    )
    op.create_table(
        "nota_custo",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("nota_id", sa.Integer(), nullable=False),
        sa.Column("ordem", sa.Integer(), nullable=False),
        sa.Column("rotulo", sa.Text(), nullable=False),
        sa.Column("categoria", sa.String(length=20), nullable=False),
        sa.Column("grupo", sa.String(length=10), nullable=False),
        sa.Column("valor", sa.Text(), nullable=False),
        sa.ForeignKeyConstraint(
            ["nota_id"], ["nota_corretagem.id"], name=op.f("fk_nota_custo_nota_id_nota_corretagem")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_nota_custo")),
    )
    with op.batch_alter_table("nota_custo", schema=None) as batch_op:
        batch_op.create_index(batch_op.f("ix_nota_custo_nota_id"), ["nota_id"], unique=False)

    op.create_table(
        "operacao",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("nota_id", sa.Integer(), nullable=False),
        sa.Column("ordem", sa.Integer(), nullable=False),
        sa.Column("categoria", sa.String(length=10), nullable=False),
        sa.Column("cv", sa.String(length=1), nullable=False),
        sa.Column("papel", sa.String(length=20), nullable=False),
        sa.Column("ticker_impresso", sa.String(length=20), nullable=False),
        sa.Column("ativo_id", sa.Integer(), nullable=False),
        sa.Column("mercado", sa.String(length=20), nullable=False),
        sa.Column("preco", sa.Text(), nullable=False),
        sa.Column("quantidade", sa.Text(), nullable=False),
        sa.Column("total", sa.Text(), nullable=False),
        sa.Column("ativo_objeto", sa.String(length=20), nullable=True),
        sa.Column("prazo_termo_dias", sa.Integer(), nullable=True),
        sa.ForeignKeyConstraint(
            ["ativo_id"], ["ativo.id"], name=op.f("fk_operacao_ativo_id_ativo")
        ),
        sa.ForeignKeyConstraint(
            ["nota_id"], ["nota_corretagem.id"], name=op.f("fk_operacao_nota_id_nota_corretagem")
        ),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_operacao")),
    )
    with op.batch_alter_table("operacao", schema=None) as batch_op:
        batch_op.create_index(batch_op.f("ix_operacao_ativo_id"), ["ativo_id"], unique=False)
        batch_op.create_index(batch_op.f("ix_operacao_nota_id"), ["nota_id"], unique=False)


def downgrade() -> None:
    with op.batch_alter_table("operacao", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_operacao_nota_id"))
        batch_op.drop_index(batch_op.f("ix_operacao_ativo_id"))

    op.drop_table("operacao")
    with op.batch_alter_table("nota_custo", schema=None) as batch_op:
        batch_op.drop_index(batch_op.f("ix_nota_custo_nota_id"))

    op.drop_table("nota_custo")
    op.drop_table("nota_corretagem")
    with op.batch_alter_table("ticker_historico", schema=None) as batch_op:
        batch_op.drop_index("ix_ticker_historico_ticker")

    op.drop_table("ticker_historico")
    op.drop_table("conta")
    op.drop_table("titular")
    op.drop_table("serie_opcao")
    op.drop_table("corretora")
    op.drop_table("ativo")
    op.drop_table("arquivo_importado")
