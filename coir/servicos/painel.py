"""Números do painel: o ano do titular resumido a partir da apuração do motor e das DARFs."""

from dataclasses import dataclass
from decimal import Decimal

from coir.domain.fiscal import Categoria, MesApurado
from coir.servicos.darf import LinhaDarf, Situacao

ZERO = Decimal(0)


@dataclass(frozen=True)
class ResumoAno:
    resultado: dict[Categoria, Decimal]  # resultado tributável do ano (sem o isento), antes
    # da compensação de prejuízos
    imposto_devido: Decimal  # depois do IRRF
    darfs_pagas: int
    darfs_emitidas: int  # com vencimento (fora as acumuladas e o mês em apuração)
    prejuizo_a_compensar: dict[Categoria, Decimal]  # saldo no último mês apurado (<= 0)


def resumo(meses: list[MesApurado], darfs: list[LinhaDarf]) -> ResumoAno:
    resultado = {c: ZERO for c in Categoria}
    for m in meses:
        for r in m.categorias:
            resultado[r.categoria] += r.resultado
    ultimo = meses[-1].categorias if meses else ()
    emitidas = [x for x in darfs if x.situacao not in (Situacao.ACUMULADA, Situacao.EM_APURACAO)]
    return ResumoAno(
        resultado=resultado,
        imposto_devido=sum((m.imposto_devido for m in meses), ZERO),
        darfs_pagas=sum(1 for x in emitidas if x.situacao is Situacao.PAGA),
        darfs_emitidas=len(emitidas),
        prejuizo_a_compensar={r.categoria: r.prejuizo_a_compensar for r in ultimo},
    )


def da_vez(darfs: list[LinhaDarf]) -> LinhaDarf | None:
    """A DARF a mostrar primeiro: a mais antiga em atraso; senão a próxima a vencer."""
    atrasadas = [x for x in darfs if x.situacao is Situacao.EM_ATRASO]
    a_pagar = [x for x in darfs if x.situacao is Situacao.A_PAGAR]
    for grupo in (atrasadas, a_pagar):
        if grupo:
            return min(grupo, key=lambda x: (x.ano, x.mes))
    return None
