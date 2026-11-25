"""O app do Typer, as opções e as funções de apoio comuns aos comandos (pasta, banco, data,
tabelas)."""

import logging
from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir import __version__
from coir.config import VAR_DADOS, dados_dir

if TYPE_CHECKING:
    from datetime import datetime

    from coir.domain.parametros import ReferenciasFiscais
    from coir.domain.referencias import Referencias


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


OpcaoBanco = typer.Option(None, help="Arquivo do banco (padrão: coir.db na pasta de dados).")


OpcaoCaminho = typer.Option(None, help="Padrão: dentro da pasta de dados (COIR_DADOS).")


OpcaoTitularUnico = typer.Option(
    False, help="Trata todos os CPFs da planilha como um único titular."
)


def _origem(origem: Path | None) -> Path:
    from coir.gabarito.leitor import localizar_planilha

    return origem or localizar_planilha(dados_dir(exigir=True) / "_gabarito")


def _destino(destino: Path | None) -> Path:
    return destino or dados_dir(exigir=True) / "gabarito"


def _banco(banco: Path | None) -> Path:
    from coir.db.sessao import NOME_BANCO

    return banco or dados_dir(exigir=True) / NOME_BANCO


def _agora() -> "datetime":
    from datetime import datetime

    return datetime.now().replace(microsecond=0)


# ------------------------------------------------------------------------------ eventos (F4)
ARQUIVO_EVENTOS = "eventos.csv"


# ------------------------------------------------------------------------------ apuração (F4)
def _mes(texto: str) -> tuple[int, int]:
    import re

    m = re.fullmatch(r"(\d{4})-(\d{2})", texto)
    if m is None or not 1 <= int(m[2]) <= 12:
        raise typer.BadParameter("use AAAA-MM")
    return int(m[1]), int(m[2])


def _pasta_refs(dados: Path | None = None) -> Path:
    """Pasta das tabelas em uso (o pacote baixado em <dados>/referencias, se mais novo), com a
    fonte da bolsa ligada ou não pela configuração (``fontes.b3``)."""
    from coir import referencias
    from coir.config import modo_fonte_b3

    referencias.definir_fonte_b3(modo_fonte_b3())
    return referencias.pasta_efetiva(dados or dados_dir())


def _refs() -> "Referencias":
    """Tabelas de referência: o pacote baixado em <dados>/referencias, se mais novo."""
    from coir import referencias

    return referencias.carregar(_pasta_refs())


def _fiscais() -> "ReferenciasFiscais":
    from coir import referencias

    return referencias.carregar_fiscais(_pasta_refs())
