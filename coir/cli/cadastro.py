"""Comandos do banco (migrar) e dos titulares."""

from pathlib import Path
from typing import TYPE_CHECKING

import typer

if TYPE_CHECKING:
    pass
from coir.cli.comum import OpcaoBanco, _agora, _banco, app

# ------------------------------------------------------------------------------- banco (F3)
banco_app = typer.Typer(
    help="Banco de dados local (coir.db na pasta de dados).", no_args_is_help=True
)
app.add_typer(banco_app, name="banco")


titular_app = typer.Typer(help="Titulares (CPF) do banco.", no_args_is_help=True)
app.add_typer(titular_app, name="titular")


@banco_app.command("migrar")
def banco_migrar(banco: Path | None = OpcaoBanco) -> None:
    """Cria o banco ou o leva à última versão do esquema (backup antes)."""
    from coir.db.sessao import migrar

    caminho = _banco(banco)
    migrar(caminho, _agora())
    print(f"Banco atualizado: {caminho}")


@titular_app.command("adicionar")
def titular_adicionar(
    cpf: str = typer.Argument(..., help="CPF do titular (com ou sem máscara)."),
    nome: str = typer.Option(..., help="Nome do titular."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Cadastra um titular (CPF) no banco."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar, migrar

    caminho = _banco(banco)
    migrar(caminho, _agora())
    with conectar(caminho) as fabrica, fabrica.begin() as sessao:
        titular = repo.adicionar_titular(sessao, cpf, nome)
        print(f"Titular cadastrado: {titular.nome} (id {titular.id})")


@titular_app.command("listar")
def titular_listar(banco: Path | None = OpcaoBanco) -> None:
    """Lista os titulares do banco."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar

    with conectar(_banco(banco)) as fabrica, fabrica() as sessao:
        for t in repo.titulares(sessao):
            print(f"{t.id}: {t.nome}")


@titular_app.command("renomear")
def titular_renomear(
    id_titular: int = typer.Argument(..., help="Id do titular (coir titular listar)."),
    nome: str = typer.Option(..., help="Novo nome."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Troca o nome de um titular (o criado na importação usa o nome impresso na nota)."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar

    with conectar(_banco(banco)) as fabrica, fabrica.begin() as sessao:
        titular = repo.renomear_titular(sessao, id_titular, nome)
        print(f"Titular {titular.id}: {titular.nome}")
