"""Relatórios fiscais de um titular e um ano (F5), montados só a partir de `DadosRelatorio`.

Cada função devolve um `Relatorio` neutro: a exportação (xlsx, csv, pdf) e a interface (F6)
usam o mesmo objeto. Nenhum cálculo fiscal novo aqui: os valores vêm da apuração do motor
(regra 12: cada número tem origem no motor ou na nota).
"""

from coir.reports.fiscais.apuracao import (
    darfs,
    evolucao,
    irpf_mensal,
    irrf_exclusiva,
    prejuizos,
    resumo_fiscal,
)
from coir.reports.fiscais.comum import GRUPOS, MESES, NOME_CATEGORIA, NOME_GRUPO, _total
from coir.reports.fiscais.lista import ABAS, TODOS, relatorios
from coir.reports.fiscais.operacoes import (
    conferencia,
    itens_finalizadas,
    livro_notas,
    operacoes_finalizadas,
)
from coir.reports.fiscais.patrimonio import NOME_PROVENTO, bens_e_direitos, carteira, proventos

__all__ = [
    "ABAS",
    "GRUPOS",
    "MESES",
    "NOME_CATEGORIA",
    "NOME_GRUPO",
    "NOME_PROVENTO",
    "TODOS",
    "_total",
    "bens_e_direitos",
    "carteira",
    "conferencia",
    "darfs",
    "evolucao",
    "irpf_mensal",
    "irrf_exclusiva",
    "itens_finalizadas",
    "livro_notas",
    "operacoes_finalizadas",
    "prejuizos",
    "proventos",
    "relatorios",
    "resumo_fiscal",
]
