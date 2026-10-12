"""Leitura do .xlsb (valores em cache, sem Excel). Única parte do gabarito que conhece o pyxlsb."""

from collections.abc import Iterable
from pathlib import Path

from pyxlsb import open_workbook

from coir.erros import GabaritoError

Linhas = list[list[object]]


def ler_abas(caminho: Path, nomes: Iterable[str]) -> dict[str, Linhas]:
    """Lê as abas pedidas como listas de valores (None para célula vazia)."""
    if not caminho.is_file():
        raise GabaritoError(f"Planilha-gabarito não encontrada: {caminho}")
    abas: dict[str, Linhas] = {}
    with open_workbook(str(caminho)) as livro:
        existentes = set(livro.sheets)
        for nome in nomes:
            if nome not in existentes:
                raise GabaritoError(f"Aba ausente na planilha-gabarito: {nome}")
            with livro.get_sheet(nome) as aba:
                abas[nome] = [[celula.v for celula in linha] for linha in aba.rows()]
    return abas
