"""Apuração mensal por titular: categorias, isenção, compensação, IRRF e DARF.

Fontes (alíquotas e limites vêm de data/parametros_fiscais.csv, com a fonte de cada um):
- Alíquotas: 15% comum (Lei 11.033/2004 art. 2º II), 20% day trade e 20% FII (IN RFB
  1.585/2015, arts. 56 a 65).
- Isenção: ganho líquido em ações no mercado à vista quando as alienações do mês somam até
  R$ 20 mil (Lei 11.033/2004 art. 3º I); units, ETF, BDR, FII, opções, futuros e day trade fora.
- Prejuízo: compensa só na mesma categoria, sem prazo (IN RFB 1.585/2015); prejuízo de mês isento
  compensa se `prejuizo_mes_isento_compensavel` (Lei 8.981/1995 art. 72 §8º).
- IRRF (0,005% e 1%): deduz do imposto do mês; o saldo passa aos meses seguintes do ano; o de
  dezembro vai para a declaração anual (IN RFB 1.585/2015; decisão do usuário). A DARF usa um
  saldo único; as fichas do programa IRPF recebem o IRRF que cada uma consumiu (`fichas.py`).
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
    FichaMes,
    Grupo,
    IrrfAnual,
    MesApurado,
    Movimento,
    NotaMotor,
    OutroCusto,
    ResultadoCategoria,
    TipoAtivo,
    TributacaoExclusiva,
)
from coir.domain.parametros import Calendario, ParametrosFiscais
from coir.engine.calendario import fim_do_mes, vencimento_darf
from coir.engine.fichas import Retido, SaldoPorFicha, irrf_retido
from coir.engine.futuros import movimentos_futuros
from coir.engine.movimentos import movimentar
from coir.erros import MotorError

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
    outros_custos: Sequence[OutroCusto] = (),
) -> Apuracao:
    """Apuração de todos os titulares, mês a mês, até o mês de `ate` (visão por titular)."""
    avisos: list[str] = []
    movimentos = movimentar(notas, eventos, parametros, ate, avisos=avisos)
    movimentos += movimentos_futuros(notas, ate, avisos)
    movimentos += [m for c in outros_custos if (m := _outro_custo(c)).data <= ate]
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
    irrf = irrf_retido([n for n in notas if n.data <= ate], movimentos)
    inicio: dict[int, date] = {}
    for n in notas:
        if n.data <= ate:
            inicio[n.titular] = min(inicio.get(n.titular, n.data), n.data)
    for e in eventos:
        if e.data <= ate:
            inicio[e.titular] = min(inicio.get(e.titular, e.data), e.data)
    meses: list[MesApurado] = []
    anuais: list[IrrfAnual] = []
    for titular in sorted(inicio):
        m, a = _apurar_titular(titular, _meses(inicio[titular], ate), por_mes, irrf, pme,
                               parametros, calendario, avisos)  # fmt: skip
        meses += m
        anuais += a
    return Apuracao(movimentos=tuple(movimentos), meses=tuple(meses), exclusiva=tuple(exclusiva),
                    irrf_anual=tuple(anuais), avisos=tuple(avisos))  # fmt: skip


def _exclusiva(m: Movimento) -> TributacaoExclusiva:
    return TributacaoExclusiva(data=m.data, titular=m.titular, ativo=m.ativo or 0, ticker=m.ticker,
                               quantidade=m.quantidade, alienacao=m.valor, custo=m.custo_baixa,
                               resultado=m.resultado)  # fmt: skip


def _outro_custo(c: OutroCusto) -> Movimento:
    """Outro custo dedutível do mês como movimento (regra 12: aparece na memória do mês)."""
    if c.categoria is Categoria.FII or c.valor < 0:
        raise MotorError(f"outro custo {c.id}: categoria {c.categoria} ou valor {c.valor} inválido")
    return Movimento(
        data=fim_do_mes(c.ano, c.mes), titular=c.titular, conta=None, corretora=None, ativo=None,
        ticker="", origem=f"custo:{c.id}",
        regra="outros custos dedutíveis lançados pelo usuário: reduzem o resultado do mês"
        + (f" ({c.descricao})" if c.descricao else ""),
        categoria=c.categoria, grupo=Grupo.A_VISTA, tipo=TipoAtivo.ACAO, cv=None,
        quantidade=ZERO, valor=ZERO, custos=c.valor, resultado=-c.valor,
        quantidade_apos=ZERO, custo_apos=ZERO,
    )  # fmt: skip


def _apurar_titular(
    titular: int,
    meses: Iterable[tuple[int, int]],
    por_mes: dict[tuple[int, int, int], list[Movimento]],
    irrf: dict[tuple[int, int, int], Retido],
    pme: dict[tuple[int, int, int], Decimal],
    params: ParametrosFiscais,
    cal: Calendario,
    avisos: list[str],
) -> tuple[list[MesApurado], list[IrrfAnual]]:
    prejuizo = {c: ZERO for c in Categoria}
    vendas_mes: dict[tuple[int, int], Decimal] = defaultdict(lambda: ZERO)
    for (t, ano, mes), movs in por_mes.items():
        if t == titular:
            vendas_mes[(ano, mes)] = sum((_vendas_no_limite(m, params) for m in movs), ZERO)
    saldo = SaldoPorFicha()
    acumulado_darf = ZERO
    irrf_ano = Retido()
    saida: list[MesApurado] = []
    anuais: list[IrrfAnual] = []
    pagos: set[tuple[int, int]] = set()  # meses cuja DARF foi paga (no mês ou acumulada depois)
    pendentes: list[tuple[int, int]] = []
    anos: dict[int, tuple[SaldoPorFicha, list[MesApurado]]] = {}
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
        fii = next(r.imposto for r in categorias if r.categoria is Categoria.FII)
        if mes == 1 or ano not in anos:
            saldo, irrf_ano = SaldoPorFicha(), Retido()
            anos[ano] = (saldo, [])
        saldo_anterior, saldo_dt_anterior = saldo.saldo, saldo.saldo_day_trade
        retido = irrf.get((titular, ano, mes), Retido())
        irrf_ano += retido
        saldo.reter(mes, retido)
        # o IRRF de day trade é consumido primeiro pela ficha de comuns: o saldo dele não vai para
        # a declaração anual (só restituição), o de 0,005% vai (decisão do usuário, 15/10/2026)
        deduzido = saldo.deduzir(mes, imposto - fii, fii)
        devido = imposto - deduzido
        total = devido + acumulado_darf
        pendentes.append((ano, mes))
        if total >= params.decimal("darf_minimo", dia):
            darf, novo_acumulado, venc = total, ZERO, vencimento_darf(ano, mes, cal)
            pagos.update(pendentes)
            pendentes = []
        else:
            darf, novo_acumulado, venc = ZERO, total, None
        apurado = MesApurado(
            titular=titular, ano=ano, mes=mes, categorias=tuple(categorias), imposto=imposto,
            irrf_mes=retido.day_trade + retido.lei, irrf_saldo_anterior=saldo_anterior,
            irrf_deduzido=deduzido, irrf_saldo=saldo.saldo, imposto_devido=devido,
            darf_acumulado_anterior=acumulado_darf, darf=darf,
            darf_acumulado=novo_acumulado, vencimento=venc,
            isento_pme=centavos(pme.get((titular, ano, mes), ZERO)),
            irrf_mes_day_trade=retido.day_trade, irrf_saldo_day_trade_anterior=saldo_dt_anterior,
            irrf_saldo_day_trade=saldo.saldo_day_trade,
        )  # fmt: skip
        saida.append(apurado)
        anos[ano][1].append(apurado)
        acumulado_darf = novo_acumulado
        if mes == 12:
            anuais.append(IrrfAnual(titular=titular, ano=ano, saldo=saldo.saldo,
                                    saldo_comum=saldo.saldo - saldo.saldo_day_trade,
                                    saldo_day_trade=saldo.saldo_day_trade,
                                    comum=irrf_ano.lei, day_trade=irrf_ano.day_trade))  # fmt: skip
    fichas: dict[tuple[int, int], tuple[FichaMes, ...]] = {}
    for ano, (s, ms) in anos.items():
        do_ano = s.fichas(ms, {m for a, m in pagos if a == ano}, avisos)
        fichas.update({(ano, mes): f for mes, f in do_ano.items()})
    saida = [m.model_copy(update={"fichas": fichas[(m.ano, m.mes)]}) for m in saida]
    anuais = [
        a.model_copy(update={"saldo_comum_fichas": {
            f.ficha: f.irrf_a_compensar for f in fichas[(a.ano, 12)]}})
        for a in anuais
    ]  # fmt: skip
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
