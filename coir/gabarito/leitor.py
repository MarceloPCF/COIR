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


def localizar_planilha(pasta: Path) -> Path:
    """A planilha-gabarito é o único .xlsb da pasta (o nome é livre; a v2 usa o CPF)."""
    candidatas = sorted(
        p for p in pasta.glob("*.xlsb") if p.is_file() and not p.name.startswith("~$")
    )
    if len(candidatas) != 1:
        raise GabaritoError(
            f"Esperada uma única planilha .xlsb em {pasta}, há {len(candidatas)}; "
            "indique o arquivo com --origem"
        )
    return candidatas[0]
