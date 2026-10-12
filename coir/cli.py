"""Linha de comando do COIR (Typer)."""

import logging
import sys

import typer

from coir import __version__
from coir.config import VAR_DADOS, dados_dir
from coir.erros import CoirError

app = typer.Typer(
    help="COIR v3 — Controle de Operações e Imposto de Renda (B3).",
    no_args_is_help=True,
)


@app.callback()
def principal(
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Log detalhado."),
) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


@app.command()
def versao() -> None:
    """Mostra a versão do COIR."""
    print(f"COIR {__version__}")


@app.command()
def config() -> None:
    """Mostra a pasta de dados configurada e se ela existe."""
    pasta = dados_dir()
    situacao = "existe" if pasta.is_dir() else "NÃO existe"
    print(f"Pasta de dados ({VAR_DADOS}): {pasta} [{situacao}]")


def main() -> None:
    """Ponto de entrada: erros de domínio viram mensagem curta e código de saída 1."""
    try:
        app()
    except CoirError as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
