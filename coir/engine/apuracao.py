"""Apuração mensal por titular: categorias, isenção, compensação, IRRF e DARF.

Fontes (alíquotas e limites vêm de data/parametros_fiscais.csv, com a fonte de cada um):
- Alíquotas: 15% comum (Lei 11.033/2004 art. 2º II), 20% day trade e 20% FII (IN RFB
  1.585/2015, arts. 56 a 65).
- Isenção: ganho líquido em ações no mercado à vista quando as alienações do mês somam até
  R$ 20 mil (Lei 11.033/2004 art. 3º I); units, ETF, BDR, FII, opções, futuros e day trade fora.
- Prejuízo: compensa só na mesma categoria, sem prazo (IN RFB 1.585/2015); prejuízo de mês isento
  compensa se `prejuizo_mes_isento_compensavel` (Lei 8.981/1995 art. 72 §8º).
- IRRF (0,005% e 1%): deduz do imposto do mês; o saldo passa aos meses seguintes do ano; o de
  dezembro vai para a declaração anual (IN RFB 1.585/2015; decisão do usuário).
- DARF 6015: abaixo de R$ 10 acumula, inclusive para o ano seguinte (Lei 9.430/1996 art. 68
  §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915).
"""

from collections import defaultdict
from collections.abc import Iterable, Sequence
from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

from coir.domain.fiscal import (
    ZERO,
    Apuracao,
    Categoria,
    Evento,
    Grupo,
    IrrfAnual,
    MesApurado,
    Movimento,
    NotaMotor,
    ResultadoCategoria,
    TipoAtivo,
    TributacaoExclusiva,
)
from coir.domain.parametros import Calendario, ParametrosFiscais
from coir.engine.calendario import fim_do_mes, vencimento_darf
from coir.engine.futuros import movimentos_futuros
from coir.engine.movimentos import movimentar

CENTAVO = Decimal("0.01")
ALIQUOTAS = {
    Categoria.COMUM: "aliquota_comum",
    Categoria.DAY_TRADE: "aliquota_day_trade",
    Categoria.FII: "aliquota_fii",
}


def centavos(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVO, rounding=ROUND_HALF_UP)


def truncar(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVO, rounding=ROUND_DOWN)


def _meses(inicio: date, fim: date) -> list[tuple[int, int]]:
    meses: list[tuple[int, int]] = []
    ano, mes = inicio.year, inicio.month
    while (ano, mes) <= (fim.year, fim.month):
        meses.append((ano, mes))
        ano, mes = (ano + 1, 1) if mes == 12 else (ano, mes + 1)
    return meses


def apurar(
    notas: Sequence[NotaMotor],
    eventos: Sequence[Evento],
    parametros: ParametrosFiscais,
    calendario: Calendario,
    ate: date,
) -> Apuracao:
    """Apuração de todos os titulares, mês a mês, até o mês de `ate` (visão por titular)."""
    avisos: list[str] = []
    movimentos = movimentar(notas, eventos, parametros, ate, avisos=avisos)
    movimentos += movimentos_futuros(notas, ate, avisos)
    movimentos.sort(key=lambda m: m.data)
    por_mes: dict[tuple[int, int, int], list[Movimento]] = defaultdict(list)
    exclusiva: list[TributacaoExclusiva] = []
    pme: dict[tuple[int, int, int], Decimal] = defaultdict(lambda: ZERO)
    for m in movimentos:
        if m.tipo is TipoAtivo.ETF_RF:
            if m.origem.startswith("op:") and (m.cv == "V" or m.resultado):
                exclusiva.append(_exclusiva(m))
            continue
        if m.isento_pme:
            pme[(m.titular, m.data.year, m.data.month)] += m.resultado
            continue
        por_mes[(m.titular, m.data.year, m.data.month)].append(m)
    irrf: dict[tuple[int, int, int], tuple[Decimal, Decimal]] = defaultdict(lambda: (ZERO, ZERO))
    inicio: dict[int, date] = {}
    for n in notas:
        if n.data <= ate:
            comum, dt = irrf[(n.titular, n.data.year, n.data.month)]
            irrf[(n.titular, n.data.year, n.data.month)] = (comum + n.irrf, dt + n.irrf_day_trade)
            inicio[n.titular] = min(inicio.get(n.titular, n.data), n.data)
    for e in eventos:
        if e.data <= ate:
            inicio[e.titular] = min(inicio.get(e.titular, e.data), e.data)
    meses: list[MesApurado] = []
    anuais: list[IrrfAnual] = []
    for titular in sorted(inicio):
        m, a = _apurar_titular(titular, _meses(inicio[titular], ate), por_mes, irrf, pme,
                               parametros, calendario)  # fmt: skip
        meses += m
        anuais += a
    return Apuracao(movimentos=tuple(movimentos), meses=tuple(meses), exclusiva=tuple(exclusiva),
                    irrf_anual=tuple(anuais), avisos=tuple(avisos))  # fmt: skip


def _exclusiva(m: Movimento) -> TributacaoExclusiva:
    return TributacaoExclusiva(data=m.data, titular=m.titular, ativo=m.ativo or 0, ticker=m.ticker,
                               quantidade=m.quantidade, alienacao=m.valor, custo=m.custo_baixa,
                               resultado=m.resultado)  # fmt: skip


def _apurar_titular(
    titular: int,
    meses: Iterable[tuple[int, int]],
    por_mes: dict[tuple[int, int, int], list[Movimento]],
    irrf: dict[tuple[int, int, int], tuple[Decimal, Decimal]],
    pme: dict[tuple[int, int, int], Decimal],
    params: ParametrosFiscais,
    cal: Calendario,
) -> tuple[list[MesApurado], list[IrrfAnual]]:
    prejuizo = {c: ZERO for c in Categoria}
    vendas_mes: dict[tuple[int, int], Decimal] = defaultdict(lambda: ZERO)
    for (t, ano, mes), movs in por_mes.items():
        if t == titular:
            vendas_mes[(ano, mes)] = sum((_vendas_no_limite(m, params) for m in movs), ZERO)
    saldo_irrf = ZERO
    saldo_dt = ZERO  # parte do saldo que é IRRF de day trade (consumida primeiro)
    acumulado_darf = ZERO
    irrf_ano = (ZERO, ZERO)
    saida: list[MesApurado] = []
    anuais: list[IrrfAnual] = []
    for ano, mes in meses:
        dia = fim_do_mes(ano, mes)
        movs = por_mes.get((titular, ano, mes), [])
        categorias: list[ResultadoCategoria] = []
        for cat in Categoria:
            r = _categoria(cat, [m for m in movs if m.categoria is cat], prejuizo[cat], params, dia,
                           vendas_mes)  # fmt: skip
            prejuizo[cat] = r.prejuizo_a_compensar
            categorias.append(r)
        imposto = sum((r.imposto for r in categorias), ZERO)
        if mes == 1:
            saldo_irrf, saldo_dt, irrf_ano = ZERO, ZERO, (ZERO, ZERO)
        comum, dt = irrf[(titular, ano, mes)]
        irrf_ano = (irrf_ano[0] + comum, irrf_ano[1] + dt)
        disponivel = saldo_irrf + comum + dt
        deduzido = min(imposto, disponivel)
        # o IRRF de day trade é consumido primeiro: o saldo dele não vai para a declaração anual
        # (só restituição), o de 0,005% vai (decisão do usuário, 15/10/2026)
        saldo_dt_anterior = saldo_dt
        saldo_dt = max(saldo_dt + dt - deduzido, ZERO)
        devido = imposto - deduzido
        total = devido + acumulado_darf
        if total >= params.decimal("darf_minimo", dia):
            darf, novo_acumulado, venc = total, ZERO, vencimento_darf(ano, mes, cal)
        else:
            darf, novo_acumulado, venc = ZERO, total, None
        saida.append(
            MesApurado(
                titular=titular, ano=ano, mes=mes, categorias=tuple(categorias), imposto=imposto,
                irrf_mes=comum + dt, irrf_saldo_anterior=saldo_irrf, irrf_deduzido=deduzido,
                irrf_saldo=disponivel - deduzido, imposto_devido=devido,
                darf_acumulado_anterior=acumulado_darf, darf=darf,
                darf_acumulado=novo_acumulado, vencimento=venc,
                isento_pme=centavos(pme.get((titular, ano, mes), ZERO)),
                irrf_mes_day_trade=dt, irrf_saldo_day_trade_anterior=saldo_dt_anterior,
                irrf_saldo_day_trade=saldo_dt,
            )
        )  # fmt: skip
        saldo_irrf, acumulado_darf = disponivel - deduzido, novo_acumulado
        if mes == 12:
            anuais.append(IrrfAnual(titular=titular, ano=ano, saldo=saldo_irrf,
                                    saldo_comum=saldo_irrf - saldo_dt, saldo_day_trade=saldo_dt,
                                    comum=irrf_ano[0], day_trade=irrf_ano[1]))  # fmt: skip
    return saida, anuais


def _vendas_no_limite(m: Movimento, params: ParametrosFiscais) -> Decimal:
    """Valor bruto que entra na soma do limite de R$ 20 mil (units só com o parâmetro)."""
    if m.categoria is not Categoria.COMUM:
        return ZERO
    if m.tipo is TipoAtivo.UNIT and not params.booleano("units_na_isencao", m.data):
        return ZERO
    return m.venda_acao_a_vista


def _categoria(
    cat: Categoria,
    movs: Sequence[Movimento],
    prejuizo_anterior: Decimal,
    params: ParametrosFiscais,
    dia: date,
    vendas_mes: dict[tuple[int, int], Decimal],
) -> ResultadoCategoria:
    por_grupo = {g: ZERO for g in Grupo}
    for m in movs:
        por_grupo[m.grupo] += m.resultado
    total = sum(por_grupo.values(), ZERO)
    vendas = ZERO
    isento = ZERO
    if cat is Categoria.COMUM:
        vendas = vendas_mes.get((dia.year, dia.month), ZERO)
        units = params.booleano("units_na_isencao", dia)
        # resultado isentável por mês de referência do limite (o da venda, na venda a descoberto)
        por_ref: dict[tuple[int, int], Decimal] = defaultdict(lambda: ZERO)
        for m in movs:
            if m.isentavel and (units or m.tipo is not TipoAtivo.UNIT):
                ref = m.mes_limite or m.data
                por_ref[(ref.year, ref.month)] += m.resultado
        compensa = params.booleano("prejuizo_mes_isento_compensavel", dia)
        for (ano, mes), isentavel in por_ref.items():
            limite = params.decimal("limite_isencao_acoes", fim_do_mes(ano, mes))
            if vendas_mes.get((ano, mes), ZERO) <= limite and (isentavel > 0 or not compensa):
                isento += isentavel
                por_grupo[Grupo.A_VISTA] -= isentavel
                total -= isentavel
    resultado = centavos(total)
    saldo = resultado + prejuizo_anterior
    aliquota = params.decimal(ALIQUOTAS[cat], dia)
    base = saldo if saldo > 0 else ZERO
    return ResultadoCategoria(
        categoria=cat, por_grupo={g: centavos(v) for g, v in por_grupo.items()},
        resultado=resultado, vendas_acoes=centavos(vendas), isento=centavos(isento),
        prejuizo_anterior=prejuizo_anterior, base=base,
        prejuizo_a_compensar=saldo if saldo < 0 else ZERO, aliquota=aliquota,
        imposto=truncar(base * aliquota),
    )  # fmt: skip
