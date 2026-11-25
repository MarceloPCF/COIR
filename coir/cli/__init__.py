"""Linha de comando do COIR (Typer).

Os comandos ficam em arquivos por assunto e se registram no ``app`` ao serem importados, na ordem
abaixo (a mesma da ajuda): comum (versao, config), gabarito, notas, cadastro (banco, titular),
importacao (importar, eventos, b3) e apuracao (apurar, carteira, exportar, ui).
"""

import sys

# a ordem dos imports é a do registro dos comandos (a da ajuda): não reordenar
# isort: off
from coir.cli.comum import app
from coir.cli import gabarito  # noqa: F401
from coir.cli import notas  # noqa: F401
from coir.cli import cadastro  # noqa: F401
from coir.cli import importacao  # noqa: F401
from coir.cli import apuracao  # noqa: F401

# isort: on
from coir.erros import CoirError

__all__ = ["app", "main"]


def main() -> None:
    """Ponto de entrada: erros de domínio viram mensagem curta e código de saída 1."""
    try:
        app()
    except CoirError as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        sys.exit(1)
