"""Linhas da nota BM&F: tipo do negócio, contrato e câmbio implícito (conferência da nota).

- Tipo: só NORMAL, DAY TRADE, AJUPOS e taxa de permanência; qualquer outro (entrega física,
  exercício de opção sobre futuro) falha alto: a nota não entra calada como ajuste.
- Contrato: raiz de 3 letras + mês e ano (WINM24) ou PPPP (permanência). Opção sobre futuro
  (ex.: ICFK26C300) fica fora do escopo e falha alto (decisão do usuário, 19/10/2026).
- Câmbio implícito: na nota, o valor de cada negócio é o ajuste ao preço de ajuste do dia (PA),
  ``valor = (PA - preço) x quantidade x multiplicador x câmbio`` (compra; a venda tem o sinal
  trocado). O PA aparece na linha AJUPOS do mesmo contrato; sem ela, os negócios do dia que se
  fecham (quantidade líquida zero) dispensam o PA. Contrato cotado em R$: o câmbio tem de dar 1, o
  que confere o multiplicador nota a nota; em US$ (café, soja CME): o mesmo câmbio em todas as
  linhas. Diferença acima do arredondamento é erro de conciliação (decisão do usuário, 19/10/2026).
"""

import re
from collections import defaultdict
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import TipoAjuste
from coir.domain.nota import NegocioBMF, NotaExtraida
from coir.domain.referencias import ContratoFuturo, Referencias
from coir.erros import ConciliacaoError, ExtracaoError, ReferenciaError
from coir.parsing.numeros import compactar

RE_CONTRATO = re.compile(r"^[A-Z]{3}(?:[FGHJKMNQUVXZ]\d{2}|PPPP)$")
TIPOS = {"NORMAL": TipoAjuste.NORMAL, "DAYTRADE": TipoAjuste.DAY_TRADE,
         "AJUPOS": TipoAjuste.AJUPOS}  # fmt: skip
CENTAVO = Decimal("0.01")
UM = Decimal(1)


def tipo(n: NegocioBMF) -> TipoAjuste:
    """Tipo do negócio impresso; desconhecido falha alto."""
    t = compactar(n.tipo).upper()
    if t in TIPOS:
        return TIPOS[t]
    if "PERMAN" in t:
        return TipoAjuste.OUTRO
    raise ExtracaoError(f"{n.mercadoria}: tipo de negócio BM&F {n.tipo!r} fora do escopo "
                        "(ex.: entrega física, exercício de opção sobre futuro)")  # fmt: skip


def contrato(n: NegocioBMF, refs: Referencias, dia: date) -> ContratoFuturo | None:
    """Contrato da linha; None só para mercadoria sem cadastro (quem chama decide se falha)."""
    if not RE_CONTRATO.match(n.mercadoria):
        raise ExtracaoError(f"{n.mercadoria}: código fora do formato de contrato futuro (opção "
                            "sobre futuro e outros derivativos estão fora do escopo)")  # fmt: skip
    return next((c for c in refs.contratos if c.raiz == n.mercadoria[:3] and c.vigente(dia)), None)


def _assinado(n: NegocioBMF) -> Decimal:
    return n.valor if n.dc == "C" else -n.valor


def _sinal(n: NegocioBMF) -> Decimal:
    return UM if n.cv == "C" else -UM


def cambios(nota: NotaExtraida, refs: Referencias) -> dict[str, Decimal | None]:
    """Câmbio implícito por contrato da nota (None: não dá para calcular). Confere cada linha e
    levanta ConciliacaoError se o valor impresso não fecha com o preço de ajuste."""
    por_contrato: dict[str, list[NegocioBMF]] = defaultdict(list)
    for n in nota.negocios_bmf:
        if tipo(n) in (TipoAjuste.NORMAL, TipoAjuste.DAY_TRADE, TipoAjuste.AJUPOS):
            por_contrato[n.mercadoria].append(n)
    saida: dict[str, Decimal | None] = {}
    for mercadoria, linhas in por_contrato.items():
        c = contrato(linhas[0], refs, nota.data_pregao)
        if c is None:
            raise ReferenciaError(f"contrato futuro {mercadoria} sem cadastro")
        saida[mercadoria] = _cambio(nota, mercadoria, linhas, c)
    return saida


def _cambio(nota: NotaExtraida, mercadoria: str, linhas: list[NegocioBMF],
            c: ContratoFuturo) -> Decimal | None:  # fmt: skip
    negocios = [n for n in linhas if tipo(n) is not TipoAjuste.AJUPOS]
    ajustes = {n.preco for n in linhas if tipo(n) is TipoAjuste.AJUPOS}
    if len(ajustes) > 1:
        raise ConciliacaoError(f"{mercadoria}: preço de ajuste", "um só", sorted(ajustes),
                               f"{nota.arquivo}: ")  # fmt: skip
    pares: list[tuple[Decimal, Decimal]] = []  # (base sem câmbio, valor impresso)
    if ajustes:
        pa = next(iter(ajustes))
        pares = [(_sinal(n) * (pa - n.preco) * abs(n.quantidade) * c.multiplicador,
                  _assinado(n)) for n in negocios]  # fmt: skip
    elif negocios and sum((_sinal(n) * abs(n.quantidade) for n in negocios), Decimal(0)) == 0:
        base = -sum((_sinal(n) * n.preco * abs(n.quantidade) for n in negocios), Decimal(0))
        pares = [(base * c.multiplicador, sum((_assinado(n) for n in negocios), Decimal(0)))]
    referencia = max(pares, key=lambda p: abs(p[0]), default=None)
    if referencia is None or referencia[0] == 0:
        return UM if c.moeda == "BRL" else None
    cambio = UM if c.moeda == "BRL" else referencia[1] / referencia[0]
    taxa = f", câmbio {cambio.quantize(Decimal('0.0001'))}" if c.moeda == "USD" else ""
    for base, valor in pares:
        esperado = (base * cambio).quantize(CENTAVO)
        if abs(esperado - valor) > CENTAVO:
            raise ConciliacaoError(
                f"{mercadoria}: valor pelo preço de ajuste (multiplicador {c.multiplicador}{taxa})",
                esperado, valor, f"{nota.arquivo}: ",
            )  # fmt: skip
    return cambio


def conversao_do_volume(nota: NotaExtraida, refs: Referencias) -> dict[str, Decimal]:
    """Fator que leva o volume de cada contrato (preço x quantidade x multiplicador) a reais, para
    ratear os custos da nota. Nota de uma moeda só: tudo 1 (a proporção não muda). Nota que mistura
    contratos em R$ e em US$: o câmbio implícito; sem ele, falha alto."""
    taxas = cambios(nota, refs)
    moedas = {mercadoria: _moeda(mercadoria, refs, nota.data_pregao) for mercadoria in taxas}
    if len(set(moedas.values())) <= 1:
        return dict.fromkeys(taxas, UM)
    saida: dict[str, Decimal] = {}
    for mercadoria, taxa in taxas.items():
        if taxa is None:
            raise ConciliacaoError(f"{mercadoria}: câmbio implícito para ratear os custos",
                                   "calculável (linha de ajuste ou negócios fechados no dia)",
                                   "indisponível", f"{nota.arquivo}: ")  # fmt: skip
        saida[mercadoria] = taxa
    return saida


def _moeda(mercadoria: str, refs: Referencias, dia: date) -> str:
    c = next((c for c in refs.contratos if c.raiz == mercadoria[:3] and c.vigente(dia)), None)
    return c.moeda if c else "BRL"
