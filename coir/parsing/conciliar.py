"""Conciliação da nota (CLAUDE.md, regra 4): os valores lidos têm de fechar entre si.

Bovespa: negócios == resumo dos negócios; valor líquido das operações == vendas - compras. BM&F:
negócios == valor dos negócios e ajuste de day trade. Essas somas são exatas: diferença aí é leitura
errada, e a nota não entra.

Os subtotais de custo e o líquido são outra história. Cada valor da nota é arredondado à parte pela
corretora, então a soma do que está impresso pode não bater ao centavo com o total impresso, sem
erro nenhum. A diferença vira a linha de custo "Ajuste de arredondamento" e a nota fecha: o que foi
pago é o líquido impresso. Só o que passar do que o arredondamento explica (meio centavo por parcela
somada) consome a tolerância da nota (R$ 0,05, com aviso); acima dela, a nota não entra.
"""

from collections.abc import Iterable
from dataclasses import dataclass
from decimal import Decimal
from itertools import product

from coir.domain.nota import (
    ROTULO_AJUSTE,
    CategoriaCusto,
    Custo,
    GrupoCusto,
    Montante,
    NegocioBovespa,
    NotaExtraida,
    Segmento,
)
from coir.erros import ConciliacaoError
from coir.parsing.observacoes import MARCAS_COM_AVISO

TOLERANCIA_NOTA = Decimal("0.05")  # além do arredondamento explicado, por nota
MEIO_CENTAVO = Decimal("0.005")  # erro máximo de cada valor impresso, arredondado ao centavo
ZERO = Decimal(0)


@dataclass(frozen=True)
class _Diferenca:
    """Quanto falta somar aos custos (débito) para ``campo`` fechar com o impresso."""

    campo: str
    valor: Decimal
    parcelas: int  # valores impressos somados: cada um pode vir com até meio centavo de erro
    grupo: GrupoCusto | None  # None: só confere (o valor do resumo já é o que vale)
    esperado: Decimal
    lido: Decimal

    @property
    def nao_explicado(self) -> Decimal:
        return max(abs(self.valor) - self.parcelas * MEIO_CENTAVO, ZERO)


def _contexto(nota: NotaExtraida) -> str:
    return f"{nota.arquivo}: "


def _exigir(campo: str, esperado: Decimal, lido: Decimal, nota: NotaExtraida) -> None:
    if esperado != lido:
        raise ConciliacaoError(campo, esperado, lido, _contexto(nota))


def _debito(m: Montante) -> Decimal:
    """Subtotal de custos: sem indicador D/C impresso, é débito."""
    return -m.valor if m.dc == "C" else m.valor


def _impressos(custos: Iterable[Custo]) -> tuple[Custo, ...]:
    """Os custos da nota sem o ajuste de uma conciliação anterior: conciliar de novo dá o mesmo."""
    return tuple(c for c in custos if c.rotulo != ROTULO_AJUSTE)


def _soma(custos: Iterable[Custo], grupo: GrupoCusto) -> Decimal:
    return sum((c.valor for c in custos if c.grupo == grupo), ZERO)


def _parcelas(custos: Iterable[Custo], grupo: GrupoCusto) -> int:
    """Itens não zerados do subtotal mais o próprio total impresso."""
    return sum(1 for c in custos if c.grupo == grupo and c.valor != 0) + 1


def _parcelas_liquido(*subtotais: Montante) -> int:
    """Subtotais não zerados mais o próprio líquido impresso."""
    return sum(1 for s in subtotais if s.valor != 0) + 1


def _outros_custos(nota: NotaExtraida) -> Decimal:
    """O que soma em ``custos_total`` fora das linhas de custo (impostos da BM&F)."""
    return nota.resumo_bmf.impostos if nota.resumo_bmf else ZERO


def _fechar(
    nota: NotaExtraida, custos: tuple[Custo, ...], diferencas: list[_Diferenca]
) -> tuple[tuple[Custo, ...], tuple[str, ...]]:
    """Custos com o ajuste de arredondamento e os avisos; acima da tolerância, a nota não entra."""
    nao_explicado = sum((d.nao_explicado for d in diferencas), ZERO)
    if nao_explicado > TOLERANCIA_NOTA:
        pior = max(diferencas, key=lambda d: d.nao_explicado)
        raise ConciliacaoError(pior.campo, pior.esperado, pior.lido, _contexto(nota))
    ajustes: dict[GrupoCusto, Decimal] = {}
    for d in diferencas:
        if d.grupo is not None and d.valor != 0:
            ajustes[d.grupo] = ajustes.get(d.grupo, ZERO) + d.valor
    avisos: tuple[str, ...] = ()
    # custo negativo não existe (IN RFB 1.585/2015 art. 56 §3º): o ajuste para no custo zero
    sem_ajuste = sum((c.valor for c in custos), ZERO) + _outros_custos(nota)
    falta = -(sem_ajuste + sum(ajustes.values(), ZERO))
    if falta > 0:
        for g in list(ajustes):
            corte = min(falta, max(-ajustes[g], ZERO))
            ajustes[g] += corte
            falta -= corte
        avisos = (
            f"nota {nota.numero} de {nota.arquivo}: o ajuste de arredondamento deixaria o custo "
            "negativo; ficou em custo zero",
        )
    extras = tuple(
        Custo(rotulo=ROTULO_AJUSTE, categoria=CategoriaCusto.OUTROS, grupo=g, valor=v)
        for g, v in ajustes.items() if v != 0
    )  # fmt: skip
    if nao_explicado > 0:
        campos = ", ".join(d.campo for d in diferencas if d.nao_explicado > 0)
        avisos = (
            *avisos,
            f"nota {nota.numero} de {nota.arquivo}: diferença de R$ {nao_explicado} em {campos}, "
            "além do arredondamento e dentro da tolerância de R$ 0,05; a nota entrou, com o valor "
            "impresso no líquido. Confira o total da nota",
        )
    return custos + extras, avisos


def conciliar(nota: NotaExtraida) -> NotaExtraida:
    """Devolve a nota conciliada (tratamento do IRRF descoberto e ajuste de arredondamento) ou
    levanta ConciliacaoError."""
    if nota.segmento is Segmento.BOVESPA:
        return _bovespa(nota)
    return _bmf(nota)


def _avisos_day_trade(nota: NotaExtraida) -> list[str]:
    """Marca D só na compra ou só na venda: o outro lado pode estar em outra nota do dia. A
    apuração confere as quantidades do dia e avisa se não fecharem (``coir.engine.movimentos``)."""
    lados: dict[str, set[str]] = {}
    for n in nota.negocios_bovespa:
        if n.day_trade:
            lados.setdefault(n.especificacao, set()).add(n.cv)
    return [
        f"nota {nota.numero} de {nota.arquivo}: day trade de {papel} marcado só na "
        f"{'compra' if cvs == {'C'} else 'venda'}; o outro lado deve estar em outra nota do dia"
        for papel, cvs in lados.items() if cvs != {"C", "V"}
    ]  # fmt: skip


def _avisos_marcas(nota: NotaExtraida, negocios: Iterable[NegocioBovespa]) -> list[str]:
    vistas: dict[str, str] = {}
    for n in negocios:
        for marca in n.observacao:
            if marca in MARCAS_COM_AVISO:
                vistas.setdefault(marca, n.especificacao)
    return [
        f"nota {nota.numero} de {nota.arquivo}: marca {m} ({MARCAS_COM_AVISO[m]}) em {papel}: o "
        "negócio entrou como operação comum; confira se esse tratamento está certo"
        for m, papel in vistas.items()
    ]


def _bovespa(nota: NotaExtraida) -> NotaExtraida:
    r = nota.resumo_bovespa
    assert r is not None
    custos = _impressos(r.custos)
    # termo: soma em "Operações à termo" e no valor das operações, mas não no líquido da nota
    a_termo = [n for n in nota.negocios_bovespa if n.mercado == "TERMO"]
    demais = [n for n in nota.negocios_bovespa if n.mercado != "TERMO"]
    termo = sum((n.valor for n in a_termo), ZERO)
    compras = sum((n.valor for n in demais if n.cv == "C"), ZERO)
    vendas = sum((n.valor for n in demais if n.cv == "V"), ZERO)
    _exigir("operações à termo", r.termo, termo, nota)
    _exigir("valor das operações", r.valor_operacoes, compras + vendas + termo, nota)
    _exigir("compras (resumo dos negócios)", r.compras, compras, nota)
    _exigir("vendas (resumo dos negócios)", r.vendas, vendas, nota)
    _exigir(
        "valor líquido das operações", vendas - compras, r.valor_liquido_operacoes.assinado, nota
    )
    cblc = r.valor_liquido_operacoes.assinado - _soma(custos, "cblc")
    lido_bolsa, lido_despesas = _debito(r.total_bovespa), _debito(r.total_despesas)
    bolsa, despesas = _soma(custos, "bolsa"), _soma(custos, "despesas")
    # IRRF somado nas despesas (a corretora o inclui no total) ou fora delas
    # com IRRF de centavos, a diferença pode ser só arredondamento: aí precisa bater exato
    folga = _parcelas(custos, "despesas") * MEIO_CENTAVO
    irrf_nas_despesas = (
        r.irrf > 0
        and despesas != lido_despesas
        and abs(despesas + r.irrf - lido_despesas) < abs(despesas - lido_despesas)
        and (r.irrf > 2 * folga or despesas + r.irrf == lido_despesas)
    )
    com_despesas = despesas + (r.irrf if irrf_nas_despesas else ZERO)
    base = r.total_cblc.assinado - lido_bolsa - lido_despesas
    # IRRF fora das despesas: algumas corretoras descontam do líquido, outras só no mês (BTG, Ágora)
    tentativas = [
        (com_irrf, base - (r.irrf if com_irrf and not irrf_nas_despesas else ZERO)
         - (r.irrf_day_trade if com_dt else ZERO) - r.liquido.assinado)
        for com_irrf, com_dt in product((False, True), repeat=2)
    ]  # fmt: skip
    com_irrf, sobra = min(tentativas, key=lambda t: abs(t[1]))  # se empatar, a primeira
    diferencas = [
        _Diferenca("Total CBLC", cblc - r.total_cblc.assinado, _parcelas(custos, "cblc"), "cblc",
                   cblc, r.total_cblc.assinado),
        _Diferenca("Total Bovespa / Soma", lido_bolsa - bolsa, _parcelas(custos, "bolsa"), "bolsa",
                   bolsa, lido_bolsa),
        _Diferenca("Total de despesas", lido_despesas - com_despesas,
                   _parcelas(custos, "despesas") + int(irrf_nas_despesas), "despesas",
                   com_despesas, lido_despesas),
        _Diferenca("líquido", sobra, _parcelas_liquido(r.total_cblc, r.total_bovespa,
                                                       r.total_despesas), "despesas",
                   r.liquido.assinado + sobra, r.liquido.assinado),
    ]  # fmt: skip
    finais, avisos = _fechar(nota, custos, diferencas)
    no_liquido = irrf_nas_despesas or (com_irrf and r.irrf > 0)
    return nota.model_copy(
        update={
            "resumo_bovespa": r.model_copy(
                update={"custos": finais, "irrf_no_liquido": no_liquido}
            ),
            "avisos": (*avisos, *_avisos_day_trade(nota),
                       *_avisos_marcas(nota, nota.negocios_bovespa)),
        }
    )  # fmt: skip


def _bmf(nota: NotaExtraida) -> NotaExtraida:
    r = nota.resumo_bmf
    assert r is not None
    custos = _impressos(r.custos)
    ajuste_dt = sum(
        (n.valor if n.dc == "C" else -n.valor for n in nota.negocios_bmf if n.day_trade), ZERO
    )
    _exigir("ajuste day trade", ajuste_dt, r.ajuste_day_trade.assinado, nota)
    soma = sum((n.valor if n.dc == "C" else -n.valor for n in nota.negocios_bmf), ZERO)
    _exigir("valor dos negócios", soma, r.valor_negocios.assinado, nota)
    ajustes = r.ajuste_posicao.assinado + r.ajuste_day_trade.assinado
    _exigir("ajuste de posição + ajuste day trade", r.valor_negocios.assinado, ajustes, nota)
    taxa_linhas = sum((n.taxa_operacional for n in nota.negocios_bmf), ZERO)
    taxa_resumo = sum((c.valor for c in custos if c.categoria.value == "corretagem"), ZERO)
    soma_custos = sum((c.valor for c in custos), ZERO)
    base = r.ajuste_day_trade.assinado + r.ajuste_posicao.assinado - r.total_custos - r.impostos
    sobra = min(
        (base - extra - r.liquido.assinado
         for extra in (r.irrf_day_trade + r.irrf, r.irrf_day_trade, r.irrf, ZERO)),
        key=abs,
    )  # fmt: skip
    com_linhas = sum(1 for n in nota.negocios_bmf if n.taxa_operacional != 0)
    diferencas = [
        _Diferenca("taxa operacional", taxa_resumo - taxa_linhas, com_linhas + 1, None,
                   taxa_resumo, taxa_linhas),
        _Diferenca("total de custos", r.total_custos - soma_custos, _parcelas(custos, "bmf"), "bmf",
                   soma_custos, r.total_custos),
        _Diferenca("líquido", sobra, 3, "bmf", r.liquido.assinado + sobra, r.liquido.assinado),
    ]  # fmt: skip
    finais, avisos = _fechar(nota, custos, diferencas)
    return nota.model_copy(
        update={"resumo_bmf": r.model_copy(update={"custos": finais}), "avisos": avisos}
    )
