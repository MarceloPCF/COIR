"""Conciliação da nota (CLAUDE.md, regra 4): os valores lidos têm de fechar entre si.

Bovespa: negócios == resumo dos negócios; valor líquido das operações == vendas - compras; Total
CBLC, Total Bovespa/Soma e Total de despesas == soma dos seus itens; líquido == CBLC - Bolsa -
Despesas (- IRRF quando a corretora o desconta). BM&F: negócios == valor dos negócios e ajuste de
day trade; custos == total de custos; líquido == ajustes - custos - impostos - IRRF.

Qualquer diferença levanta ConciliacaoError e a nota não entra.
"""

from collections.abc import Iterable
from decimal import Decimal
from itertools import product

from coir.domain.nota import Custo, GrupoCusto, Montante, NotaExtraida, Segmento
from coir.erros import ConciliacaoError

TOLERANCIA = Decimal(0)  # os valores impressos fecham ao centavo nas 2.032 notas reais


def _igual(a: Decimal, b: Decimal) -> bool:
    return abs(a - b) <= TOLERANCIA


def _contexto(nota: NotaExtraida) -> str:
    return f"{nota.arquivo}: "


def _exigir(campo: str, esperado: Decimal, lido: Decimal, nota: NotaExtraida) -> None:
    if not _igual(esperado, lido):
        raise ConciliacaoError(campo, esperado, lido, _contexto(nota))


def _debito(m: Montante) -> Decimal:
    """Subtotal de custos: sem indicador D/C impresso, é débito."""
    return -m.valor if m.dc == "C" else m.valor


def _soma(custos: Iterable[Custo], grupo: GrupoCusto) -> Decimal:
    return sum((c.valor for c in custos if c.grupo == grupo), Decimal(0))


def conciliar(nota: NotaExtraida) -> NotaExtraida:
    """Devolve a nota conciliada (tratamento do IRRF descoberto) ou levanta ConciliacaoError."""
    if nota.segmento is Segmento.BOVESPA:
        return _bovespa(nota)
    return _bmf(nota)


def _bovespa(nota: NotaExtraida) -> NotaExtraida:
    r = nota.resumo_bovespa
    assert r is not None
    compras = sum((n.valor for n in nota.negocios_bovespa if n.cv == "C"), Decimal(0))
    vendas = sum((n.valor for n in nota.negocios_bovespa if n.cv == "V"), Decimal(0))
    _exigir("valor das operações", r.valor_operacoes, compras + vendas, nota)
    _exigir("compras (resumo dos negócios)", r.compras, compras, nota)
    _exigir("vendas (resumo dos negócios)", r.vendas, vendas, nota)
    _exigir(
        "valor líquido das operações", vendas - compras, r.valor_liquido_operacoes.assinado, nota
    )
    lados: dict[str, set[str]] = {}
    for n in nota.negocios_bovespa:
        if n.day_trade:
            lados.setdefault(n.especificacao, set()).add(n.cv)
    for papel, cvs in lados.items():
        if cvs != {"C", "V"}:  # marca D sem compra e venda do mesmo papel: marca lida errado
            raise ConciliacaoError(
                f"day trade de {papel}", "compra e venda", "/".join(sorted(cvs)), _contexto(nota)
            )
    cblc = r.valor_liquido_operacoes.assinado - _soma(r.custos, "cblc")
    _exigir("Total CBLC", cblc, r.total_cblc.assinado, nota)
    _exigir("Total Bovespa / Soma", _soma(r.custos, "bolsa"), _debito(r.total_bovespa), nota)
    despesas = _soma(r.custos, "despesas")
    lido_despesas = _debito(r.total_despesas)
    irrf_nas_despesas = not _igual(despesas, lido_despesas) and _igual(
        despesas + r.irrf, lido_despesas
    )
    _exigir(
        "Total de despesas", despesas + (r.irrf if irrf_nas_despesas else 0), lido_despesas, nota
    )
    base = cblc - _debito(r.total_bovespa) - lido_despesas
    # IRRF fora das despesas: algumas corretoras descontam do líquido, outras só no mês (BTG, Ágora)
    for com_irrf, com_dt in product((False, True), repeat=2):
        extra = (r.irrf if com_irrf and not irrf_nas_despesas else 0) + (
            r.irrf_day_trade if com_dt else 0
        )
        if _igual(base - extra, r.liquido.assinado):
            no_liquido = irrf_nas_despesas or (com_irrf and r.irrf > 0)
            return nota.model_copy(
                update={"resumo_bovespa": r.model_copy(update={"irrf_no_liquido": no_liquido})}
            )
    raise ConciliacaoError("líquido", base, r.liquido.assinado, _contexto(nota))


def _bmf(nota: NotaExtraida) -> NotaExtraida:
    r = nota.resumo_bmf
    assert r is not None
    ajuste_dt = sum(
        (n.valor if n.dc == "C" else -n.valor for n in nota.negocios_bmf if n.day_trade), Decimal(0)
    )
    _exigir("ajuste day trade", ajuste_dt, r.ajuste_day_trade.assinado, nota)
    soma = sum((n.valor if n.dc == "C" else -n.valor for n in nota.negocios_bmf), Decimal(0))
    _exigir("valor dos negócios", soma, r.valor_negocios.assinado, nota)
    ajustes = r.ajuste_posicao.assinado + r.ajuste_day_trade.assinado
    _exigir("ajuste de posição + ajuste day trade", r.valor_negocios.assinado, ajustes, nota)
    taxa_linhas = sum((n.taxa_operacional for n in nota.negocios_bmf), Decimal(0))
    taxa_resumo = sum((c.valor for c in r.custos if c.categoria.value == "corretagem"), Decimal(0))
    _exigir("taxa operacional", taxa_resumo, taxa_linhas, nota)
    _exigir("total de custos", sum((c.valor for c in r.custos), Decimal(0)), r.total_custos, nota)
    base = r.ajuste_day_trade.assinado + r.ajuste_posicao.assinado - r.total_custos - r.impostos
    for extra in (r.irrf_day_trade + r.irrf, r.irrf_day_trade, r.irrf, Decimal(0)):
        if _igual(base - extra, r.liquido.assinado):
            return nota
    raise ConciliacaoError("líquido", base, r.liquido.assinado, _contexto(nota))
