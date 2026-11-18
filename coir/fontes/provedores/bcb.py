"""Banco Central do Brasil: Selic mensal (SGS 4390), para os juros do DARF em atraso."""

from decimal import Decimal
from itertools import pairwise
from pathlib import Path

from coir.erros import ConfiguracaoError
from coir.fontes.leitura import linhas

SELIC = "selic_mensal.csv"


def selic(pasta: Path) -> dict[tuple[int, int], Decimal]:
    """Selic mensal (% a.m.): meses completos, contíguos, sem repetição, 2 casas."""
    arquivo = pasta / SELIC
    taxas: dict[tuple[int, int], Decimal] = {}
    for x in linhas(arquivo):
        try:
            mes = (int(x["mes"][:4]), int(x["mes"][5:7]))
            taxa = Decimal(x["taxa"])
        except (ValueError, ArithmeticError) as erro:
            raise ConfiguracaoError(f"{arquivo.name}: linha ilegível {x}") from erro
        if mes in taxas or taxa < 0 or taxa != taxa.quantize(Decimal("0.01")):
            raise ConfiguracaoError(f"{arquivo.name}: {x['mes']} repetido ou taxa inválida")
        taxas[mes] = taxa
    for (a1, m1), (a2, m2) in pairwise(sorted(taxas)):
        if (a2, m2) != ((a1 + 1, 1) if m1 == 12 else (a1, m1 + 1)):
            raise ConfiguracaoError(f"{arquivo.name}: falta a Selic depois de {m1:02d}/{a1}")
    return taxas
