"""Pagamentos de DARF (F6).

Revision ID: 0003
Revises: 0002
"""

from collections.abc import Sequence

import sqlalchemy as sa
from alembic import op

revision: str = "0003"
down_revision: str | None = "0002"
branch_labels: str | Sequence[str] | None = None
depends_on: str | Sequence[str] | None = None


def upgrade() -> None:
    op.create_table(
        "darf_pagamento",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("titular_id", sa.Integer(), nullable=False),
        sa.Column("ano", sa.Integer(), nullable=False),
        sa.Column("mes", sa.Integer(), nullable=False),
        sa.Column("data_pagamento", sa.Date(), nullable=False),
        sa.Column("valor_pago", sa.Text(), nullable=False),
        sa.Column("comprovante", sa.Text(), nullable=True),
        sa.Column("observacao", sa.Text(), nullable=False),
        sa.Column("registrado_em", sa.DateTime(), nullable=False),
        sa.ForeignKeyConstraint(["titular_id"], ["titular.id"],
                                name=op.f("fk_darf_pagamento_titular_id_titular")),
        sa.PrimaryKeyConstraint("id", name=op.f("pk_darf_pagamento")),
        sa.UniqueConstraint("titular_id", "ano", "mes",
                            name=op.f("uq_darf_pagamento_titular_id")),
    )  # fmt: skip


def downgrade() -> None:
    op.drop_table("darf_pagamento")
