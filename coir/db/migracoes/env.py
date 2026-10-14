"""Ambiente Alembic configurado por código (sem alembic.ini): a conexão vem de
``config.attributes["connection"]`` (ver coir.db.sessao.migrar)."""

from alembic import context

from coir.db.modelos import Base
from coir.erros import BancoError

conexao = context.config.attributes.get("connection")
if conexao is None:
    raise BancoError("migrações do COIR rodam por `coir banco migrar` (coir.db.sessao.migrar)")
context.configure(
    connection=conexao,
    target_metadata=Base.metadata,
    render_as_batch=True,  # SQLite: ALTER TABLE por recriação da tabela
    compare_type=True,
)
with context.begin_transaction():
    context.run_migrations()
