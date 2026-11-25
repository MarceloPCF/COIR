"""A lista dos relatórios fiscais, a aba de cada um no xlsx e a geração dos pedidos."""

from collections.abc import Callable, Iterable
from dataclasses import replace

from coir.domain.relatorio import DadosRelatorio
from coir.reports.fiscais.apuracao import (
    darfs,
    evolucao,
    irpf_mensal,
    irrf_exclusiva,
    prejuizos,
    resumo_fiscal,
)
from coir.reports.fiscais.operacoes import conferencia, livro_notas, operacoes_finalizadas
from coir.reports.fiscais.patrimonio import bens_e_direitos, carteira, proventos
from coir.reports.irpf import declaracao
from coir.reports.modelo import (
    Relatorio,
)

TODOS: dict[str, Callable[[DadosRelatorio], Relatorio]] = {
    "resumo-fiscal": resumo_fiscal,
    "darfs": darfs,
    "irpf-mensal": irpf_mensal,
    "irrf-e-exclusiva": irrf_exclusiva,
    "operacoes-finalizadas": operacoes_finalizadas,
    "evolucao": evolucao,
    "prejuizos": prejuizos,
    "conferencia-pm": conferencia,
    "carteira": carteira,
    "bens-e-direitos": bens_e_direitos,
    "livro-de-notas": livro_notas,
    "proventos": proventos,
    "declaracao-irpf": declaracao,
}


ABAS = {
    "resumo-fiscal": "Resumo fiscal", "darfs": "DARFs", "irpf-mensal": "IRPF mês a mês",
    "irrf-e-exclusiva": "IRRF e exclusiva", "operacoes-finalizadas": "Operações finalizadas",
    "evolucao": "Evolução", "prejuizos": "Prejuízos", "conferencia-pm": "Conferência do PM",
    "carteira": "Carteira", "bens-e-direitos": "Bens e Direitos",
    "livro-de-notas": "Livro de notas", "proventos": "Proventos",
    "declaracao-irpf": "Declaração IRPF",
}  # fmt: skip


def relatorios(d: DadosRelatorio, chaves: Iterable[str] | None = None) -> list[Relatorio]:
    escolhidos = list(TODOS) if chaves is None else list(chaves)
    return [replace(TODOS[c](d), aba=ABAS[c]) for c in escolhidos]
