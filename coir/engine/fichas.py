"""IRRF por ficha do programa IRPF (comuns/day trade e FII), com a DARF no saldo único.

A DARF mensal deduz o IRRF de um saldo único (Lei 11.033/2004 art. 2º §7º; IN RFB 1.585/2015
arts. 63 §8º e 65 §8º; decisão do usuário de 16/10/2026: separar só na declaração). O programa
IRPF tem uma ficha para comuns/day trade e outra para FII, cada uma com o próprio IR fonte e
saldo. Para as fichas fecharem com a DARF sem usar o mesmo IRRF duas vezes, cada IRRF retido
vira um lote (mês, natureza, ficha de origem) e o valor deduzido do saldo único é atribuído aos
lotes que cada ficha consumiu:

- comuns/day trade consome primeiro o day trade, depois o 0,005% retido nela e por último o
  0,005% retido em vendas de FII; FII consome o 0,005% dela, depois o de comuns e, só se ainda
  faltar, o day trade (o programa não tem essa linha na ficha FII: vira aviso);
- na ficha, o "no mês" de cada mês é o que ela consumiu dos lotes daquele mês mais o que sobrou
  dos lotes retidos nela; os saldos das fichas somam o saldo único. Por isso o "no mês" de um mês
  pode mudar quando a apuração avança (um lote retido em março e consumido em maio entra em março
  na ficha que o consumiu).

Repartição do 0,005% de uma nota que mistura FII e outros ativos: proporcional às vendas comuns
(`Movimento.base_irrf`: sem day trade, venda por exercício, PME nem ETF de renda fixa; Lei
11.033/2004 art. 2º §§1º e 2º) de cada ficha, por
titular, corretora, conta e mês (a unidade da dispensa de R$ 1, §5º); o centavo residual fica em
comuns. IRRF de nota sem operações (BM&F) fica em comuns.
"""

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Ficha,
    FichaMes,
    MesApurado,
    Movimento,
    NotaMotor,
)
from coir.erros import MotorError

CENTAVO = Decimal("0.01")
Chave = tuple[int, int, int]  # titular, ano, mês


@dataclass(frozen=True)
class Retido:
    """IRRF retido num mês: day trade (ficha de comuns) e 0,005% repartido por ficha."""

    day_trade: Decimal = ZERO
    comuns: Decimal = ZERO
    fii: Decimal = ZERO

    def __add__(self, o: "Retido") -> "Retido":
        return Retido(self.day_trade + o.day_trade, self.comuns + o.comuns, self.fii + o.fii)

    @property
    def lei(self) -> Decimal:
        return self.comuns + self.fii


def irrf_retido(notas: Iterable[NotaMotor], movimentos: Iterable[Movimento]) -> dict[Chave, Retido]:
    """IRRF das notas por titular e mês, com o 0,005% repartido entre as fichas."""
    vendas: dict[tuple[int, str, int, int, int], list[Decimal]] = defaultdict(lambda: [ZERO, ZERO])
    for m in movimentos:
        if m.base_irrf and m.corretora is not None and m.conta is not None:
            chave = (m.titular, m.corretora, m.conta, m.data.year, m.data.month)
            vendas[chave][m.categoria is Categoria.FII] += m.base_irrf
    retido_unidade: dict[tuple[int, str, int, int, int], Decimal] = defaultdict(lambda: ZERO)
    saida: dict[Chave, Retido] = defaultdict(Retido)
    for n in notas:
        mes = (n.titular, n.data.year, n.data.month)
        saida[mes] += Retido(day_trade=n.irrf_day_trade)
        if n.operacoes:
            retido_unidade[(n.titular, n.corretora, n.conta, n.data.year, n.data.month)] += n.irrf
        else:
            saida[mes] += Retido(comuns=n.irrf)
    for unidade, valor in retido_unidade.items():
        outros, fii = vendas.get(unidade, (ZERO, ZERO))
        parte_fii = ZERO
        if fii and valor:
            parte_fii = (valor * fii / (outros + fii)).quantize(CENTAVO, rounding=ROUND_HALF_UP)
        saida[unidade[0], unidade[3], unidade[4]] += Retido(comuns=valor - parte_fii, fii=parte_fii)
    return dict(saida)


@dataclass
class _Lote:
    mes: int
    day_trade: bool  # natureza: 1% de day trade ou 0,005%
    origem: Ficha
    restante: Decimal
    usado: dict[tuple[Ficha, int], Decimal] = field(default_factory=dict[tuple[Ficha, int],
                                                                          Decimal])  # fmt: skip

    def consumir(self, ficha: Ficha, mes: int, falta: Decimal) -> Decimal:
        valor = min(falta, self.restante)
        if valor > 0:
            self.restante -= valor
            self.usado[(ficha, mes)] = self.usado.get((ficha, mes), ZERO) + valor
        return valor


class SaldoPorFicha:
    """Saldo único de IRRF de um titular num ano, com o consumo atribuído por ficha."""

    def __init__(self) -> None:
        self.lotes: list[_Lote] = []

    def reter(self, mes: int, r: Retido) -> None:
        partes = ((True, Ficha.COMUNS, r.day_trade), (False, Ficha.COMUNS, r.comuns),
                  (False, Ficha.FII, r.fii))  # fmt: skip
        for dt, origem, valor in partes:
            if valor:
                self.lotes.append(_Lote(mes, dt, origem, valor))

    def _pegar(self, ficha: Ficha, mes: int, falta: Decimal, dt: bool, origem: Ficha) -> Decimal:
        usado = ZERO
        for lote in self.lotes:
            if falta - usado <= 0:
                break
            if (lote.day_trade, lote.origem) == (dt, origem):
                usado += lote.consumir(ficha, mes, falta - usado)
        return usado

    def deduzir(self, mes: int, imposto_comuns: Decimal, imposto_fii: Decimal) -> Decimal:
        """Deduz do imposto do mês (o total é min(imposto, saldo), como no saldo único)."""
        c, f = Ficha.COMUNS, Ficha.FII
        falta = {c: imposto_comuns, f: imposto_fii}
        for ficha, dt, origem in ((c, True, c), (c, False, c), (f, False, f), (c, False, f),
                                  (f, False, c), (f, True, c)):  # fmt: skip
            falta[ficha] -= self._pegar(ficha, mes, falta[ficha], dt, origem)
        return imposto_comuns + imposto_fii - falta[c] - falta[f]

    @property
    def saldo(self) -> Decimal:
        return sum((lote.restante for lote in self.lotes), ZERO)

    @property
    def saldo_day_trade(self) -> Decimal:
        return sum((lote.restante for lote in self.lotes if lote.day_trade), ZERO)

    def fichas(self, meses: Sequence[MesApurado], pagos: set[int],
               avisos: list[str]) -> dict[int, tuple[FichaMes, ...]]:  # fmt: skip
        """Fichas de cada mês do ano (`pagos`: meses cuja DARF já foi paga)."""
        saida: dict[int, list[FichaMes]] = defaultdict(list)
        for ficha in Ficha:
            anterior = {True: ZERO, False: ZERO}
            for m in meses:
                no_mes = {True: ZERO, False: ZERO}
                usado = {True: ZERO, False: ZERO}
                for lote in self.lotes:
                    if lote.mes == m.mes:
                        no_mes[lote.day_trade] += sum(
                            (v for (fi, _), v in lote.usado.items() if fi is ficha), ZERO
                        ) + (lote.restante if lote.origem is ficha else ZERO)
                    usado[lote.day_trade] += lote.usado.get((ficha, m.mes), ZERO)
                compensar = {dt: anterior[dt] + no_mes[dt] - usado[dt] for dt in (True, False)}
                if min(compensar.values()) < 0:
                    raise MotorError(f"titular {m.titular}: IRRF da ficha {ficha} em "
                                     f"{m.mes:02d}/{m.ano} negativo")  # fmt: skip
                if ficha is Ficha.FII and usado[True]:
                    avisos.append(
                        f"titular {m.titular}, {m.mes:02d}/{m.ano}: {usado[True]} de IRRF de day "
                        "trade deduzido do imposto de FII (a ficha FII do programa IRPF não tem "
                        "essa linha)"
                    )
                imposto = _imposto(m, ficha)
                a_pagar = imposto - usado[True] - usado[False]
                saida[m.mes].append(FichaMes(
                    ficha=ficha, imposto=imposto,
                    irrf_day_trade_mes=no_mes[True], irrf_day_trade_anterior=anterior[True],
                    irrf_day_trade_a_compensar=compensar[True],
                    irrf_mes=no_mes[False], irrf_anterior=anterior[False],
                    irrf_a_compensar=compensar[False],
                    a_pagar=a_pagar, pago=a_pagar if m.mes in pagos else ZERO,
                ))  # fmt: skip
                anterior = compensar
        return {mes: tuple(f) for mes, f in saida.items()}


def _imposto(m: MesApurado, ficha: Ficha) -> Decimal:
    cats = (Categoria.FII,) if ficha is Ficha.FII else (Categoria.COMUM, Categoria.DAY_TRADE)
    return sum((c.imposto for c in m.categorias if c.categoria in cats), ZERO)
