"""Simulação de ordens hipotéticas (tela Simulação): as regras do motor, sem regra própria.

Cada ordem digitada vira uma nota sintética, acrescentada às notas reais, e ``apurar`` roda até o
mesmo ``ate`` com e sem elas: a diferença é o efeito das ordens no imposto, no limite de isenção,
no prejuízo a compensar e na carteira. Nada é gravado.

- Uma nota por ordem, número ``~SIM-0001``...: na ordem dos negócios do dia, esses números
  ficam depois das notas reais (a simulação supõe a ordem executada depois dos negócios
  já feitos no dia; o day trade casa pela ordem dos negócios, IN RFB 1.585/2015 art. 65 §3º).
- IRRF zero nas notas sintéticas: a DARF simulada nunca fica menor que a real. A estimativa do
  IRRF que a corretora reteria (0,005% e 1% do day trade) só é exibida (``irrf_estimado``).
- Base e cenário até o mesmo ``ate``: vender em meses diferentes muda o prejuízo e o IRRF dos meses
  seguintes, e a comparação é do horizonte inteiro.
"""

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_DOWN, Decimal

from coir.domain.fiscal import (
    CV,
    ZERO,
    Apuracao,
    Categoria,
    Evento,
    MesApurado,
    Modelo,
    Movimento,
    NotaMotor,
    OperacaoMotor,
    OutroCusto,
    SaldoInicial,
    TipoAtivo,
    Visao,
)
from coir.domain.parametros import Calendario, ParametrosFiscais
from coir.engine.apuracao import apurar
from coir.engine.calendario import dia_de_pregao, fim_do_mes
from coir.engine.movimentos import movimentar
from coir.erros import ReferenciaError, SimulacaoError

CENTAVO = Decimal("0.01")
# tipos que a tela simula: os da carteira à vista com regra no motor (opções, futuros, ETF de renda
# fixa, fundos de alíquota zero e FIDC ficam fora: vencimento, ajuste e tributação próprios)
TIPOS_SIMULAVEIS = frozenset({TipoAtivo.ACAO, TipoAtivo.UNIT, TipoAtivo.ETF, TipoAtivo.BDR,
                              TipoAtivo.FII, TipoAtivo.DIREITO, TipoAtivo.FIP})  # fmt: skip
PREFIXO = "~SIM-"  # "~" ordena depois de dígitos e letras: a ordem fica depois das notas reais


class OrdemSimulada(Modelo):
    """Uma ordem digitada na tela (compra ou venda à vista)."""

    numero: int  # posição na lista da tela (1, 2, ...)
    data: date
    titular: int
    conta: int
    corretora: str
    ativo: int
    ticker: str
    tipo: TipoAtivo
    cv: CV
    quantidade: Decimal
    preco: Decimal
    custos: Decimal = ZERO  # corretagem e taxas estimadas pelo usuário (reduzem o ganho)

    @property
    def total(self) -> Decimal:
        return self.quantidade * self.preco


@dataclass(frozen=True)
class Entradas:
    """As entradas reais do motor, lidas do banco uma vez."""

    notas: Sequence[NotaMotor]
    eventos: Sequence[Evento]
    outros_custos: Sequence[OutroCusto] = ()
    saldos_iniciais: Sequence[SaldoInicial] = ()

    def ultima_nota(self, titular: int) -> date | None:
        datas = [n.data for n in self.notas if n.titular == titular]
        return max(datas) if datas else None


class ResultadoOrdem(Modelo):
    """O que o motor fez com uma ordem no cenário."""

    numero: int
    ticker: str
    cv: CV
    quantidade: Decimal
    valor: Decimal  # bruto (quantidade x preço)
    resultado: Decimal  # operação comum ou FII (venda), sem o day trade
    quantidade_comum: Decimal  # parte que ficou como operação comum (ou FII)
    quantidade_day_trade: Decimal  # parte que casou com compra ou venda do mesmo dia
    resultado_day_trade: Decimal  # do grupo de day trade em que a ordem entrou
    quantidade_apos: Decimal  # posição do titular depois da ordem
    pm_apos: Decimal


class PosicaoFinal(Modelo):
    ticker: str
    quantidade_base: Decimal
    custo_base: Decimal
    quantidade: Decimal
    custo: Decimal


class MesSimulado(Modelo):
    ano: int
    mes: int
    base: MesApurado | None
    cenario: MesApurado


class ResultadoSimulacao(Modelo):
    titular: int
    ate: date
    meses: tuple[MesSimulado, ...]  # do mês da primeira ordem até `ate`
    ordens: tuple[ResultadoOrdem, ...]
    day_trades: tuple[Movimento, ...]  # movimentos de day trade que só existem no cenário
    posicoes: tuple[PosicaoFinal, ...]
    irrf_alienacao: Decimal  # estimativa do 0,005% a mais (exibida, não deduzida)
    irrf_day_trade: Decimal  # estimativa do 1% a mais (exibida, não deduzida)
    avisos: tuple[str, ...]  # avisos do motor que só aparecem no cenário

    @property
    def irrf_estimado(self) -> Decimal:
        return self.irrf_alienacao + self.irrf_day_trade

    def total(self, campo: str, cenario: bool = True) -> Decimal:
        """Soma de um campo do mês (imposto, imposto_devido, darf) no horizonte."""
        meses = [m.cenario if cenario else m.base for m in self.meses]
        return sum((getattr(m, campo) for m in meses if m is not None), ZERO)

    def isento(self, cenario: bool = True) -> Decimal:
        meses = [m.cenario if cenario else m.base for m in self.meses]
        return sum((r.isento for m in meses if m is not None for r in m.categorias), ZERO)

    def prejuizo_final(self, cenario: bool = True) -> dict[Categoria, Decimal]:
        """Prejuízo a compensar no fim do horizonte, por categoria (positivo)."""
        ultimo = self.meses[-1].cenario if cenario else self.meses[-1].base
        if ultimo is None:
            return {c: ZERO for c in Categoria}
        return {r.categoria: -r.prejuizo_a_compensar for r in ultimo.categorias}


def ultimo_mes(cal: Calendario) -> tuple[int, int]:
    """Último mês simulável: a DARF dele vence no mês seguinte (RIR/2018 art. 915), que o
    calendário tem de cobrir; a de dezembro venceria em janeiro do ano que ele não tem."""
    return (cal.ultimo_ano, 11)


def validar(
    ordens: Sequence[OrdemSimulada], entradas: Entradas, cal: Calendario,
    isolados: Iterable[str] = (),
) -> None:  # fmt: skip
    """Ordens que o motor apuraria de outro jeito do que o usuário quis (SimulacaoError)."""
    if not ordens:
        raise SimulacaoError("nenhuma ordem para simular")
    fora = set(isolados)
    limite = ultimo_mes(cal)
    for o in ordens:
        onde = f"ordem {o.numero} ({o.ticker})"
        if o.quantidade <= 0 or o.preco <= 0 or o.custos < 0:
            raise SimulacaoError(f"{onde}: quantidade e preço têm de ser maiores que zero, e os "
                                 "custos não podem ser negativos")  # fmt: skip
        if o.tipo not in TIPOS_SIMULAVEIS:
            raise SimulacaoError(f"{onde}: {_nome_tipo(o.tipo)} fica fora da simulação")
        if o.ticker in fora:
            raise SimulacaoError(f"{onde}: o ativo está fora do cálculo por uma pendência (em "
                                 "geral falta um evento): resolva na tela Eventos")  # fmt: skip
        if (o.data.year, o.data.month) > limite:
            ano, mes = limite
            raise SimulacaoError(f"{onde}: o calendário de feriados vai até 12/{ano}, e a DARF "
                                 f"vence no mês seguinte; simule até {mes:02d}/{ano}")  # fmt: skip
        try:
            pregao = dia_de_pregao(o.data, cal)
        except ReferenciaError as erro:
            raise SimulacaoError(f"{onde}: {erro}") from erro
        if not pregao:
            raise SimulacaoError(f"{onde}: {o.data:%d/%m/%Y} não é dia de pregão")
        ultima = entradas.ultima_nota(o.titular)
        if ultima is not None and o.data < ultima:
            raise SimulacaoError(f"{onde}: a data é anterior à última nota importada "
                                 f"({ultima:%d/%m/%Y}); simule a partir dela")  # fmt: skip


def _nome_tipo(tipo: TipoAtivo) -> str:
    return {
        TipoAtivo.OPCAO: "opção", TipoAtivo.FUTURO: "contrato futuro",
        TipoAtivo.ETF_RF: "ETF de renda fixa (tributação na fonte)",
        TipoAtivo.FUNDO_INFRA: "fundo de infraestrutura (alíquota zero)",
        TipoAtivo.FIDC: "cota de FIDC", TipoAtivo.SEM_TIPO: "ativo sem tipo nas tabelas",
    }.get(tipo, tipo.value)  # fmt: skip


def notas_sinteticas(
    ordens: Sequence[OrdemSimulada], nota_id: int, operacao_id: int
) -> tuple[list[NotaMotor], dict[int, int]]:
    """Uma nota por ordem, com ids a partir dos dados (acima dos do banco). Devolve as notas e o
    mapa id da operação -> número da ordem."""
    notas: list[NotaMotor] = []
    mapa: dict[int, int] = {}
    for k, o in enumerate(sorted(ordens, key=lambda o: (o.data, o.numero))):
        op = OperacaoMotor(
            id=operacao_id + k, ordem=0, ativo=o.ativo, ticker=o.ticker, tipo=o.tipo,
            mercado="VISTA", cv=o.cv, quantidade=o.quantidade, preco=o.preco, total=o.total,
        )  # fmt: skip
        mapa[op.id] = o.numero
        notas.append(NotaMotor(
            id=nota_id + k, titular=o.titular, conta=o.conta, corretora=o.corretora,
            numero=f"{PREFIXO}{k + 1:04d}", data=o.data, custos=o.custos, irrf=ZERO,
            irrf_day_trade=ZERO, operacoes=(op,),
        ))  # fmt: skip
    return notas, mapa


def _ids(origem: str) -> list[int]:
    if not origem.startswith("op:"):
        return []
    return [int(x) for x in origem.removeprefix("op:").split(",") if x.lstrip("-").isdigit()]


def _acima_da_posicao(
    entradas: Entradas, sinteticas: Sequence[NotaMotor], mapa: dict[int, int],
    params: ParametrosFiscais, ate: date,
) -> None:  # fmt: skip
    """Venda acima da posição da conta: o motor apuraria venda a descoberto, que não é o que a
    simulação quer mostrar (parecer de 18/11/2026). Confere na visão por conta."""
    movs = movimentar([*entradas.notas, *sinteticas], entradas.eventos, params, ate, Visao.CONTA,
                      isolados=[])  # fmt: skip
    for m in movs:
        if m.cv != "V" or m.categoria is Categoria.DAY_TRADE or m.quantidade_apos >= 0:
            continue
        numeros = sorted({mapa[i] for i in _ids(m.origem) if i in mapa})
        if numeros:
            raise SimulacaoError(
                f"ordem {numeros[0]} ({m.ticker}): a venda passa da posição da conta em "
                f"{m.data:%d/%m/%Y} ({-m.quantidade_apos:f} a mais). Venda a descoberto fica "
                "fora da simulação"
            )


def simular(
    entradas: Entradas, ordens: Sequence[OrdemSimulada], params: ParametrosFiscais,
    cal: Calendario, ate: date, base: Apuracao | None = None, desde: date | None = None,
) -> ResultadoSimulacao:  # fmt: skip
    """Apura o cenário com as ordens e compara com a base, do mês de `desde` (padrão: o da
    primeira ordem) até `ate`. Cenários comparados entre si usam o mesmo `desde` e o mesmo `ate`.
    `base`: a apuração sem as ordens até o mesmo `ate` (a tela já a tem em cache)."""
    titulares = {o.titular for o in ordens}
    if len(titulares) != 1:
        raise SimulacaoError("as ordens de uma simulação têm de ser de um titular só")
    titular = titulares.pop()
    if base is None:
        base = apurar(entradas.notas, entradas.eventos, params, cal, ate, entradas.outros_custos,
                      entradas.saldos_iniciais)  # fmt: skip
    validar(ordens, entradas, cal,
            {i.ticker for i in base.isolados if i.titular == titular})  # fmt: skip
    if ate < max(o.data for o in ordens):
        raise SimulacaoError("a data final da simulação é anterior a uma ordem")
    nota_id = max((n.id for n in entradas.notas), default=0) + 1
    operacao_id = max((o.id for n in entradas.notas for o in n.operacoes), default=0) + 1
    sinteticas, mapa = notas_sinteticas(ordens, nota_id, operacao_id)
    _acima_da_posicao(entradas, sinteticas, mapa, params, ate)
    cenario = apurar([*entradas.notas, *sinteticas], entradas.eventos, params, cal, ate,
                     entradas.outros_custos, entradas.saldos_iniciais)  # fmt: skip
    inicio = min(o.data for o in ordens)
    if desde is not None:
        if desde > inicio:
            raise SimulacaoError("o início da comparação é posterior a uma ordem")
        inicio = desde
    meses_base = {(m.ano, m.mes): m for m in base.meses if m.titular == titular}
    meses = tuple(
        MesSimulado(ano=m.ano, mes=m.mes, base=meses_base.get((m.ano, m.mes)), cenario=m)
        for m in cenario.meses
        if m.titular == titular and (m.ano, m.mes) >= (inicio.year, inicio.month)
    )  # fmt: skip
    dos_sinteticos = [m for m in cenario.movimentos if any(i in mapa for i in _ids(m.origem))]
    avisos_base = set(base.avisos_de(titular))
    return ResultadoSimulacao(
        titular=titular, ate=ate, meses=meses,
        ordens=_por_ordem(ordens, dos_sinteticos, mapa),
        day_trades=tuple(m for m in dos_sinteticos if m.categoria is Categoria.DAY_TRADE),
        posicoes=_posicoes(base, cenario, titular, ordens),
        irrf_alienacao=_irrf_alienacao(base, cenario, titular, inicio, params),
        irrf_day_trade=_irrf_day_trade(base, cenario, titular, inicio, params),
        avisos=tuple(a for a in cenario.avisos_de(titular) if a not in avisos_base),
    )  # fmt: skip


def _por_ordem(
    ordens: Sequence[OrdemSimulada], movs: Sequence[Movimento], mapa: dict[int, int]
) -> tuple[ResultadoOrdem, ...]:
    saida: list[ResultadoOrdem] = []
    for o in sorted(ordens, key=lambda o: o.numero):
        meus = [m for m in movs if o.numero in {mapa.get(i) for i in _ids(m.origem)}]
        dt = [m for m in meus if m.categoria is Categoria.DAY_TRADE]
        comuns = [m for m in meus if m.categoria is not Categoria.DAY_TRADE]
        # o day trade junta as operações do dia (a origem lista todas, inclusive as que sobraram
        # como comuns): a parte da ordem é o que não ficou como comum
        q_comum = sum((m.quantidade for m in comuns), ZERO)
        q_dt = o.quantidade - q_comum if dt else ZERO
        # posição depois da ordem: a do passo comum (o day trade não muda a posição)
        ultimo = comuns[-1] if comuns else (meus[-1] if meus else None)
        saida.append(ResultadoOrdem(
            numero=o.numero, ticker=o.ticker, cv=o.cv, quantidade=o.quantidade, valor=o.total,
            resultado=sum((m.resultado for m in comuns), ZERO),
            quantidade_comum=o.quantidade - q_dt, quantidade_day_trade=q_dt,
            resultado_day_trade=sum((m.resultado for m in dt), ZERO) if q_dt else ZERO,
            quantidade_apos=ultimo.quantidade_apos if ultimo else ZERO,
            pm_apos=ultimo.pm_apos if ultimo else ZERO,
        ))  # fmt: skip
    return tuple(saida)


def _posicoes(
    base: Apuracao, cenario: Apuracao, titular: int, ordens: Sequence[OrdemSimulada]
) -> tuple[PosicaoFinal, ...]:
    def ultimas(apuracao: Apuracao) -> dict[int, Movimento]:
        saida: dict[int, Movimento] = {}
        for m in apuracao.movimentos:  # em ordem de data
            if m.titular == titular and m.ativo is not None:
                saida[m.ativo] = m
        return saida

    b, c = ultimas(base), ultimas(cenario)
    saida: list[PosicaoFinal] = []
    for ativo, ticker in sorted({(o.ativo, o.ticker) for o in ordens}, key=lambda x: x[1]):
        mb, mc = b.get(ativo), c.get(ativo)
        saida.append(PosicaoFinal(
            ticker=ticker,
            quantidade_base=mb.quantidade_apos if mb else ZERO,
            custo_base=mb.custo_apos if mb else ZERO,
            quantidade=mc.quantidade_apos if mc else ZERO, custo=mc.custo_apos if mc else ZERO,
        ))  # fmt: skip
    return tuple(saida)


def _truncar(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVO, rounding=ROUND_DOWN)


def _irrf_alienacao(
    base: Apuracao, cenario: Apuracao, titular: int, desde: date, params: ParametrosFiscais
) -> Decimal:
    """0,005% a mais por corretora e mês: a retenção sobre a soma das alienações do mês na
    corretora, dispensada quando igual ou inferior a R$ 1,00 (Lei 11.033/2004 art. 2º §§1º, 4º e
    5º; IN RFB 1.585/2015 art. 63). Estimativa: o cálculo usa o IRRF impresso nas notas."""

    def retido(apuracao: Apuracao) -> dict[tuple[str, int, int], Decimal]:
        vendas: dict[tuple[str, int, int], Decimal] = defaultdict(lambda: ZERO)
        premios: dict[tuple[str, date], Decimal] = defaultdict(lambda: ZERO)
        for m in apuracao.movimentos:
            # a retenção e a dispensa são do mês inteiro na corretora, com as vendas reais dele
            if (m.titular != titular or m.corretora is None or not m.base_irrf
                    or (m.data.year, m.data.month) < (desde.year, desde.month)):  # fmt: skip
                continue
            if m.tipo is TipoAtivo.OPCAO:  # a soma dos prêmios do dia, se positiva (art. 63 II)
                premios[(m.corretora, m.data)] += m.base_irrf
            else:
                vendas[(m.corretora, m.data.year, m.data.month)] += m.base_irrf
        for (corretora, dia), soma in premios.items():
            vendas[(corretora, dia.year, dia.month)] += max(soma, ZERO)
        saida: dict[tuple[str, int, int], Decimal] = {}
        for (corretora, ano, mes), soma in vendas.items():
            dia = fim_do_mes(ano, mes)
            imposto = soma * params.decimal("aliquota_irrf_alienacao", dia)
            dispensa = params.decimal("irrf_alienacao_dispensa", dia)
            saida[(corretora, ano, mes)] = _truncar(imposto) if imposto > dispensa else ZERO
        return saida

    b, c = retido(base), retido(cenario)
    return sum((max(v - b.get(k, ZERO), ZERO) for k, v in c.items()), ZERO)


def _irrf_day_trade(
    base: Apuracao, cenario: Apuracao, titular: int, desde: date, params: ParametrosFiscais
) -> Decimal:
    """1% a mais sobre o resultado positivo de day trade do dia por corretora, compensadas as
    perdas do dia (Lei 9.959/2000 art. 8º; IN RFB 1.585/2015 art. 65 caput e §4º). Estimativa."""

    def retido(apuracao: Apuracao) -> dict[tuple[str | None, date], Decimal]:
        dia_a_dia: dict[tuple[str | None, date], Decimal] = defaultdict(lambda: ZERO)
        for m in apuracao.movimentos:
            if m.titular == titular and m.categoria is Categoria.DAY_TRADE and m.data >= desde:
                dia_a_dia[(m.corretora, m.data)] += m.resultado
        return {
            k: _truncar(max(v, ZERO) * params.decimal("aliquota_irrf_day_trade", k[1]))
            for k, v in dia_a_dia.items()
        }  # fmt: skip

    b, c = retido(base), retido(cenario)
    return sum((max(v - b.get(k, ZERO), ZERO) for k, v in c.items()), ZERO)
