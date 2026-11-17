"""Leitura das tabelas CSV de referência (texto, sem conversão)."""

import csv
from pathlib import Path

from coir.erros import ConfiguracaoError
from coir.parsing.numeros import compactar


def linhas(caminho: Path) -> list[dict[str, str]]:
    """Linhas da tabela, com os campos sem espaços nas pontas; tabela ausente falha alto."""
    if not caminho.is_file():
        raise ConfiguracaoError(f"tabela de referência ausente: {caminho.name}")
    with caminho.open(encoding="utf-8", newline="") as arquivo:
        return [
            {k: (v or "").strip() for k, v in linha.items()} for linha in csv.DictReader(arquivo)
        ]


def opcionais(caminho: Path) -> list[dict[str, str]]:
    """Linhas da tabela que pode faltar (pacote antigo ou fonte desligada): vazio sem ela."""
    return linhas(caminho) if caminho.is_file() else []


def mapa(caminho: Path, chave: str, valor: str, *, compacta: bool = False) -> dict[str, str]:
    """Chave -> valor da tabela; a mesma chave com dois valores falha alto. ``compacta`` tira
    espaços e pontuação da chave (nome de pregão como a nota imprime)."""
    saida: dict[str, str] = {}
    for linha in linhas(caminho):
        k = compactar(linha[chave]) if compacta else linha[chave]
        if k in saida and saida[k] != linha[valor]:
            raise ConfiguracaoError(f"{caminho.name}: {linha[chave]!r} com dois valores")
        saida[k] = linha[valor]
    return saida


def mapa_opcional(caminho: Path, chave: str, valor: str, *, compacta: bool = False
                  ) -> dict[str, str]:  # fmt: skip
    return mapa(caminho, chave, valor, compacta=compacta) if caminho.is_file() else {}
