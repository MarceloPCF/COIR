"""Simulação de ordens (tela Simulação): o motor apura as ordens digitadas como notas a mais, sem
regra própria, e compara com a base até o mesmo mês. Números conferidos à mão, sem custos e sem
IRRF nas ordens (casos S1 a S9 do parecer fiscal de 18/11/2026).

Fontes: Lei 11.033/2004 art. 2º II (15%) e §§1º, 4º e 5º (IRRF de 0,005% e dispensa mensal até
R$ 1,00) e art. 3º I (limite de R$ 20.000,00 das alienações de ações no mês); IN RFB 1.585/2015
art. 59 I e §2º (soma do limite e isenção), art. 63 (IRRF de 0,005%), art. 65 §§1º a 4º (day trade
na mesma instituição, pela ordem dos negócios; 1% sobre o resultado do dia), art. 37 §2º (prejuízo
só na mesma categoria); Lei 9.959/2000 art. 8º (day trade); Lei 9.430/1996 art. 68 §1º (DARF
abaixo de R$ 10,00 acumula); RIR/2018 art. 915 (vencimento no último dia útil do mês seguinte).
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, TipoAtivo
from coir.engine.simulacao import Entradas, OrdemSimulada, ResultadoSimulacao, simular
from coir.erros import SimulacaoError
from tests.fiscal_cases.apoio import REFS, Livro, parametros

D = Decimal
COMPRA = date(2026, 10, 5)
NOV = date(2026, 11, 23)  # 20/11 é feriado
DEZ = date(2026, 12, 14)
FIM_NOV = date(2026, 11, 30)
FIM_DEZ = date(2026, 12, 31)
ULTIMO = REFS.calendario.ultimo_ano  # a DARF de dezembro venceria fora do calendário
ESCADA = (("10.40", 1), ("10.80", 2), ("11.20", 3), ("11.60", 4))


def _livro() -> Livro:
    """Carteira base: 10.000 XXXX3 (ação) a PM 10,00 na corretora XP."""
    livro = Livro()
    livro.nota(COMPRA, livro.op("C", "XXXX3", 10000, "10"))
    return livro


def _ordem(livro: Livro, numero: int, ticker: str, quantidade: int, preco: str, dia: date = NOV,
           *, cv: str = "V", tipo: TipoAtivo | None = None, corretora: str = "XP",
           conta: int = 1, custos: str = "0") -> OrdemSimulada:  # fmt: skip
    return OrdemSimulada(
        numero=numero, data=dia, titular=1, conta=conta, corretora=corretora,
        ativo=livro.ativo(ticker), ticker=ticker, tipo=tipo or REFS.tipos.tipo(ticker),
        cv=cv, quantidade=D(quantidade), preco=D(preco), custos=D(custos),  # type: ignore[arg-type]
    )  # fmt: skip


def _simular(livro: Livro, ordens: list[OrdemSimulada], ate: date = FIM_NOV,
             **troca: str) -> ResultadoSimulacao:  # fmt: skip
    params = parametros(**troca) if troca else REFS.parametros
    return simular(Entradas(livro.notas, livro.eventos), ordens, params, REFS.calendario, ate)


def _mes(r: ResultadoSimulacao, ano: int, mes: int) -> tuple[Decimal, Decimal, date | None]:
    m = next(x.cenario for x in r.meses if (x.ano, x.mes) == (ano, mes))
    return m.imposto, m.darf, m.vencimento


def _comum(r: ResultadoSimulacao, ano: int, mes: int, cat: Categoria = Categoria.COMUM):
    m = next(x.cenario for x in r.meses if (x.ano, x.mes) == (ano, mes))
    return next(c for c in m.categorias if c.categoria is cat)


def test_s1_escada_de_quatro_ordens_todas_executadas() -> None:
    """Vendas 110.000,00 > limite: ganho 10.000,00 x 15% = 1.500,00, DARF em 30/12/2026 (31/12 não
    é dia útil bancário). Preço médio da escada 11,00; o PM da carteira não muda nas vendas.
    IRRF estimado: 110.000,00 x 0,005% = 5,50 (não deduzido)."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p) for p, n in ESCADA])
    assert _mes(r, 2026, 11) == (D("1500.00"), D("1500.00"), date(2026, 12, 30))
    assert [o.resultado for o in r.ordens] == [D(1000), D(2000), D(3000), D(4000)]
    assert all(o.pm_apos == D(10) for o in r.ordens[:3])
    assert r.ordens[-1].quantidade_apos == 0
    assert r.irrf_alienacao == D("5.50")
    assert r.irrf_day_trade == 0
    assert r.total("darf", cenario=False) == 0
    assert r.posicoes[0].quantidade_base == D(10000) and r.posicoes[0].quantidade == 0


def test_s2_execucao_parcial_so_as_duas_primeiras() -> None:
    """Vendas 53.000,00 > limite: ganho 3.000,00, imposto 450,00. Ficam 5.000 a 10,00."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p) for p, n in ESCADA[:2]])
    assert _mes(r, 2026, 11)[:2] == (D("450.00"), D("450.00"))
    assert (r.posicoes[0].quantidade, r.posicoes[0].custo) == (D(5000), D(50000))
    assert r.irrf_alienacao == D("2.65")


def test_s3_fronteira_do_limite() -> None:
    """1.923 a 10,40 = 19.999,20 (igual ou inferior ao limite): ganho 769,20 isento, IRRF 0,99996
    dispensado. 1.924 a 10,40 = 20.009,60: ganho 769,60 tributado (115,44); IRRF 1,00 retido."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1923, "10.40")])
    assert _comum(r, 2026, 11).isento == D("769.20")
    assert _mes(r, 2026, 11)[:2] == (0, 0)
    assert r.irrf_alienacao == 0
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1924, "10.40")])
    assert _mes(r, 2026, 11)[:2] == (D("115.44"), D("115.44"))
    assert r.irrf_alienacao == D("1.00")


def test_s4_ordem_que_estoura_o_limite_tira_a_isencao_de_outra_venda_do_mes() -> None:
    """Base: venda real de 1.000 YYYY3 (PM 12,00) a 15,00 em 05/11 = 15.000,00, ganho 3.000,00
    isento. Ordem de 500 XXXX3 a 10,40 (5.200,00, ganho 200,00): vendas 20.200,00, todo o ganho
    do mês tributável, 3.200,00 x 15% = 480,00. Com 480 a 10,40 (4.992,00) tudo segue isento."""
    livro = _livro()
    livro.nota(COMPRA, livro.op("C", "YYYY3", 1000, "12"))
    livro.nota(date(2026, 11, 5), livro.op("V", "YYYY3", 1000, "15"))
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 500, "10.40")])
    assert _mes(r, 2026, 11)[:2] == (D("480.00"), D("480.00"))
    base = next(m.base for m in r.meses)
    assert base is not None and base.imposto == 0
    assert next(c.isento for c in base.categorias if c.categoria is Categoria.COMUM) == D(3000)
    assert r.isento() == 0 and r.isento(cenario=False) == D(3000)
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 480, "10.40")])
    assert _mes(r, 2026, 11)[0] == 0
    assert r.isento() == D("3192.00")


def _com_prejuizo() -> Livro:
    """Prejuízo comum de 4.000,00 em outubro: 1.000 ZZZZ3 comprados a 10,00 e vendidos a 6,00
    (perda em mês isento compensa: parâmetro prejuizo_mes_isento_compensavel)."""
    livro = _livro()
    livro.nota(COMPRA, livro.op("C", "ZZZZ3", 1000, "10"))
    livro.nota(date(2026, 10, 20), livro.op("V", "ZZZZ3", 1000, "6"))
    return livro


def test_s5_prejuizo_acumulado_consumido_e_nao_consumido_pelo_isento() -> None:
    livro = _com_prejuizo()
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p) for p, n in ESCADA])
    assert _mes(r, 2026, 11)[0] == D("900.00")  # (10.000 - 4.000) x 15%
    assert r.prejuizo_final()[Categoria.COMUM] == 0
    assert r.prejuizo_final(cenario=False)[Categoria.COMUM] == D(4000)
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p) for p, n in ESCADA[:2]])
    assert _mes(r, 2026, 11)[0] == 0
    assert r.prejuizo_final()[Categoria.COMUM] == D(1000)
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1900, "10.40")])  # 19.760,00: isento
    assert _comum(r, 2026, 11).isento == D(760)
    assert r.prejuizo_final()[Categoria.COMUM] == D(4000)
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1000, "9")])  # perda em mês isento
    assert r.prejuizo_final()[Categoria.COMUM] == D(5000)


def test_s6a_distribuir_entre_meses_abaixo_do_limite_em_cada_um() -> None:
    """1.800 a 10,40 (18.720,00) e 1.800 a 10,80 (19.440,00). No mesmo mês: 38.160,00 > limite,
    ganho 2.160,00, imposto 324,00. Em novembro e dezembro: isento nos dois meses. Mesma posição
    final (6.400 a 10,00)."""
    livro = _livro()
    primeira = _ordem(livro, 1, "XXXX3", 1800, "10.40")
    junto = _simular(livro, [primeira, _ordem(livro, 2, "XXXX3", 1800, "10.80")], FIM_DEZ)
    separado = _simular(livro, [primeira, _ordem(livro, 2, "XXXX3", 1800, "10.80", DEZ)], FIM_DEZ)
    assert junto.total("imposto") == D("324.00")
    assert separado.total("imposto") == 0
    assert separado.isento() == D(2160)
    assert [len(x.meses) for x in (junto, separado)] == [2, 2]  # o mesmo horizonte
    assert junto.posicoes == separado.posicoes
    assert separado.posicoes[0].quantidade == D(6400)


def test_s6b_distribuir_acima_do_limite_so_muda_a_data_da_darf() -> None:
    """Ordens 1 e 2 em novembro (ganho 3.000,00: 450,00, vence 30/12/2026) e 3 e 4 em dezembro
    (ganho 7.000,00: 1.050,00, vence 29/01/2027): total 1.500,00, igual ao S1."""
    livro = _livro()
    datas = (NOV, NOV, DEZ, DEZ)
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p, datas[n - 1]) for p, n in ESCADA],
                 FIM_DEZ)  # fmt: skip
    assert _mes(r, 2026, 11) == (D("450.00"), D("450.00"), date(2026, 12, 30))
    assert _mes(r, 2026, 12) == (D("1050.00"), D("1050.00"), date(2027, 1, 29))
    assert r.total("darf") == D("1500.00")


def test_s6c_darf_abaixo_do_minimo_acumula_para_o_mes_seguinte() -> None:
    """ETF (sem isenção, fora do limite), 200 cotas a PM 100,00. 100 a 100,50 em novembro (ganho
    50,00: imposto 7,50, abaixo de R$ 10,00, acumula) e 100 a 100,20 em dezembro (ganho 20,00:
    3,00; DARF de 10,50 em 29/01/2027). Tudo em novembro: 10,50 vencendo em 30/12/2026."""
    livro = _livro()
    livro.nota(COMPRA, livro.op("C", "EEEE11", 200, "100", tipo=TipoAtivo.ETF))
    etf = TipoAtivo.ETF
    r = _simular(livro, [_ordem(livro, 1, "EEEE11", 100, "100.50", tipo=etf),
                         _ordem(livro, 2, "EEEE11", 100, "100.20", DEZ, tipo=etf)],
                 FIM_DEZ)  # fmt: skip
    assert _mes(r, 2026, 11) == (D("7.50"), 0, None)
    assert _mes(r, 2026, 12) == (D("3.00"), D("10.50"), date(2027, 1, 29))
    r = _simular(livro, [_ordem(livro, 1, "EEEE11", 100, "100.50", tipo=etf),
                         _ordem(livro, 2, "EEEE11", 100, "100.20", tipo=etf)],
                 FIM_DEZ)  # fmt: skip
    assert _mes(r, 2026, 11) == (D("10.50"), D("10.50"), date(2026, 12, 30))


def test_s7_fii_fora_do_limite_e_sem_o_prejuizo_comum() -> None:
    """1.000 FFFF11 (FII) a PM 95,00 e prejuízo comum de 4.000,00. Ordens: 200 FFFF11 a 100,00
    (20.000,00) e 1.900 XXXX3 a 10,40 (19.760,00). O FII não entra no limite: ganho de ações
    760,00 isento, prejuízo comum intacto; FII 1.000,00 x 20% = 200,00. IRRF estimado na mesma
    corretora: 0,005% de 39.760,00 = 1,988 -> 1,98 (acima de R$ 1,00)."""
    livro = _com_prejuizo()
    livro.nota(COMPRA, livro.op("C", "FFFF11", 1000, "95", tipo=TipoAtivo.FII))
    r = _simular(livro, [_ordem(livro, 1, "FFFF11", 200, "100", tipo=TipoAtivo.FII),
                         _ordem(livro, 2, "XXXX3", 1900, "10.40")])  # fmt: skip
    assert _comum(r, 2026, 11).isento == D(760)
    assert _comum(r, 2026, 11, Categoria.FII).imposto == D("200.00")
    assert r.prejuizo_final()[Categoria.COMUM] == D(4000)
    assert _mes(r, 2026, 11)[1] == D("200.00")
    assert r.irrf_alienacao == D("1.98")


def _com_compra_do_dia() -> Livro:
    """Nota real de 18/11/2026 na XP: compra de 1.000 XXXX3 a 11,00."""
    livro = _livro()
    livro.nota(date(2026, 11, 18), livro.op("C", "XXXX3", 1000, "11"))
    return livro


def test_s8a_venda_no_dia_de_uma_compra_na_mesma_corretora_vira_day_trade() -> None:
    """Venda de 2.500 a 11,20 na XP em 18/11: 1.000 de day trade (200,00 x 20% = 40,00) e 1.500
    comuns (ganho 1.800,00). Soma do limite 28.000,00: o ganho comum é tributado, 270,00.
    Imposto 310,00. IRRF estimado: 1% de 200,00 = 2,00; 0,005% de 16.800,00 = 0,84 dispensado."""
    livro = _com_compra_do_dia()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 2500, "11.20", date(2026, 11, 18))])
    assert _mes(r, 2026, 11)[:2] == (D("310.00"), D("310.00"))
    assert _comum(r, 2026, 11, Categoria.DAY_TRADE).imposto == D("40.00")
    o = r.ordens[0]
    assert (o.quantidade_day_trade, o.resultado_day_trade) == (D(1000), D(200))
    assert o.resultado == D(1800)
    assert len(r.day_trades) == 1
    assert r.posicoes[0].quantidade == D(8500) and r.posicoes[0].custo == D(85000)
    assert (r.irrf_day_trade, r.irrf_alienacao) == (D("2.00"), 0)


def test_s8b_mesma_venda_em_outra_corretora_nao_e_day_trade() -> None:
    """5.000 na XP e 5.000 na BTG; a compra do dia é na XP e a venda na BTG: sem day trade. PM do
    titular 111.000,00 / 11.000; custo da venda 25.227,27; ganho 2.772,73 x 15% = 415,90."""
    livro = Livro()
    livro.nota(COMPRA, livro.op("C", "XXXX3", 5000, "10"))
    livro.nota(COMPRA, livro.op("C", "XXXX3", 5000, "10"), corretora="BTG", conta=2)
    livro.nota(date(2026, 11, 18), livro.op("C", "XXXX3", 1000, "11"))
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 2500, "11.20", date(2026, 11, 18),
                                corretora="BTG", conta=2)])  # fmt: skip
    assert _mes(r, 2026, 11)[0] == D("415.90")
    assert not r.day_trades and r.ordens[0].quantidade_day_trade == 0
    assert r.posicoes[0].quantidade == D(8500)
    assert r.posicoes[0].custo.quantize(D("0.01")) == D("85772.73")


def test_s8c_day_trade_com_o_resto_isento() -> None:
    """Venda de 1.500 a 11,20 (16.800,00) na XP: day trade 1.000 (40,00); comuns 500, ganho
    600,00, isento (soma 16.800,00 no limite)."""
    livro = _com_compra_do_dia()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1500, "11.20", date(2026, 11, 18))])
    assert _mes(r, 2026, 11)[0] == D("40.00")
    assert _comum(r, 2026, 11).isento == D(600)


@pytest.mark.parametrize(("na_soma", "imposto"), [("true", "220.00"), ("false", "40.00")])
def test_s8e_day_trade_na_soma_do_limite(na_soma: str, imposto: str) -> None:
    """Venda de 2.000 a 11,20 (22.400,00): day trade 1.000 (40,00) e comuns 1.000 (ganho
    1.200,00). Com o day trade na soma (padrão), 22.400,00 > limite: +180,00; sem ele, isento.
    Com o padrão, o aviso do motor sobre o day trade na soma só aparece no cenário."""
    livro = _com_compra_do_dia()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 2000, "11.20", date(2026, 11, 18))],
                 day_trade_na_soma_limite=na_soma)  # fmt: skip
    assert _mes(r, 2026, 11)[0] == D(imposto)
    assert any("day trade" in a for a in r.avisos) is (na_soma == "true")


def test_s9_unit_fora_do_limite() -> None:
    """1.000 UUUU11 (unit, PM 14,00) a 15,00 e 900 XXXX3 a 11,00 (9.900,00): ações isentas (900,00)
    e unit tributada (1.000,00 x 15% = 150,00). Com units_na_isencao: 24.900,00 > limite, 285,00."""
    livro = _livro()
    livro.nota(COMPRA, livro.op("C", "UUUU11", 1000, "14", tipo=TipoAtivo.UNIT))
    ordens = [_ordem(livro, 1, "UUUU11", 1000, "15", tipo=TipoAtivo.UNIT),
              _ordem(livro, 2, "XXXX3", 900, "11")]  # fmt: skip
    r = _simular(livro, ordens)
    assert _mes(r, 2026, 11)[0] == D("150.00")
    assert _comum(r, 2026, 11).isento == D(900)
    assert _mes(_simular(livro, ordens, units_na_isencao="true"), 2026, 11)[0] == D("285.00")


def test_compra_hipotetica_seguida_de_venda_no_mesmo_dia_e_day_trade() -> None:
    """Compra simulada de 1.000 a 11,00 e venda simulada de 1.000 a 11,50 no mesmo dia e corretora:
    day trade de 500,00 x 20% = 100,00 (IN RFB 1.585/2015 art. 65 §§1º e 3º). IRRF estimado de 1%
    sobre 500,00 = 5,00."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1000, "11", cv="C"),
                         _ordem(livro, 2, "XXXX3", 1000, "11.50")])  # fmt: skip
    assert _comum(r, 2026, 11, Categoria.DAY_TRADE).imposto == D("100.00")
    assert r.irrf_day_trade == D("5.00")
    assert r.posicoes[0].quantidade == D(10000)


def test_custos_reduzem_o_ganho_mas_nao_a_soma_do_limite() -> None:
    """1.924 a 10,40 (20.009,60) com R$ 20,00 de custos: a soma do limite é bruta e passa dele;
    o ganho cai para 749,60 (IN RFB 1.585/2015 art. 56 §3º): 112,44."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1924, "10.40", custos="20")])
    assert _comum(r, 2026, 11).vendas_limite == D("20009.60")
    assert _mes(r, 2026, 11)[0] == D("112.44")


@pytest.mark.parametrize(
    ("ordem", "trecho"),
    [
        ({"quantidade": 10001}, "passa da posição"),
        ({"dia": date(2026, 11, 20)}, "não é dia de pregão"),
        ({"dia": date(2026, 10, 1)}, "anterior à última nota"),
        ({"dia": date(ULTIMO, 12, 1)}, f"simule até 11/{ULTIMO}"),
        ({"tipo": TipoAtivo.OPCAO}, "fica fora da simulação"),
    ],
)
def test_ordens_que_a_simulacao_recusa(ordem: dict[str, object], trecho: str) -> None:
    livro = _livro()
    quantidade = ordem.pop("quantidade", 100)
    o = _ordem(livro, 1, "XXXX3", quantidade, "10", **ordem)  # type: ignore[arg-type]
    with pytest.raises(SimulacaoError, match=trecho):
        _simular(livro, [o], date(ULTIMO, 12, 31) if o.data.year == ULTIMO else FIM_NOV)


def test_irrf_estimado_soma_as_vendas_reais_do_mes_antes_da_ordem() -> None:
    """A dispensa de R$ 1,00 é do mês inteiro na corretora (Lei 11.033/2004 art. 2º §§4º e 5º):
    venda real de 15.000,00 em 03/11 (0,75, dispensada sozinha) e ordem de 5.200,00 em 23/11
    (0,26): no mês, 20.200,00 x 0,005% = 1,01, retido."""
    livro = _livro()
    livro.nota(COMPRA, livro.op("C", "YYYY3", 1000, "12"))
    livro.nota(date(2026, 11, 3), livro.op("V", "YYYY3", 1000, "15"))
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 500, "10.40")])
    assert r.irrf_alienacao == D("1.01")


def test_duas_ordens_no_dia_so_a_primeira_casa_com_a_compra_do_dia() -> None:
    """Compra real de 1.000 a 11,00 em 18/11; ordens de venda de 1.000 a 11,20 e de 500 a 11,50
    no mesmo dia e corretora. Pela ordem dos negócios (IN RFB 1.585/2015 art. 65 §3º), a primeira
    venda casa com a compra (day trade de 200,00); a segunda é comum, pelo PM 10,00 (750,00), sem
    resultado de day trade atribuído."""
    livro = _com_compra_do_dia()
    dia = date(2026, 11, 18)
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1000, "11.20", dia),
                         _ordem(livro, 2, "XXXX3", 500, "11.50", dia)])  # fmt: skip
    primeira, segunda = r.ordens
    assert (primeira.quantidade_day_trade, primeira.resultado_day_trade) == (D(1000), D(200))
    assert primeira.quantidade_comum == 0
    assert (segunda.quantidade_day_trade, segunda.resultado_day_trade) == (0, 0)
    assert (segunda.quantidade_comum, segunda.resultado) == (D(500), D(750))
    assert segunda.quantidade_apos == D(9500) and segunda.pm_apos == D(10)


def test_compra_e_venda_no_dia_a_posicao_depois_da_ordem() -> None:
    """Na compra seguida de venda no mesmo dia, todo o volume é day trade: a posição depois de
    cada ordem continua 10.000 a PM 10,00."""
    livro = _livro()
    r = _simular(livro, [_ordem(livro, 1, "XXXX3", 1000, "11", cv="C"),
                         _ordem(livro, 2, "XXXX3", 1000, "11.50")])  # fmt: skip
    assert all(o.quantidade_day_trade == D(1000) for o in r.ordens)
    assert {(o.quantidade_apos, o.pm_apos) for o in r.ordens} == {(D(10000), D(10))}


def test_cenarios_comparados_no_mesmo_periodo() -> None:
    """Com o início comum, o cenário que só vende em dezembro também mostra novembro: os totais
    dos dois cenários cobrem os mesmos meses."""
    livro = _livro()
    ordens = [_ordem(livro, 1, "XXXX3", 2500, "10.40", DEZ)]
    params = REFS.parametros
    entradas = Entradas(livro.notas, livro.eventos)
    so_dez = simular(entradas, ordens, params, REFS.calendario, FIM_DEZ)
    desde_nov = simular(entradas, ordens, params, REFS.calendario, FIM_DEZ, desde=NOV)
    assert [(m.ano, m.mes) for m in so_dez.meses] == [(2026, 12)]
    assert [(m.ano, m.mes) for m in desde_nov.meses] == [(2026, 11), (2026, 12)]
    with pytest.raises(SimulacaoError, match="posterior a uma ordem"):
        simular(entradas, ordens, params, REFS.calendario, FIM_DEZ, desde=date(2026, 12, 15))


def test_escada_na_virada_do_ano() -> None:
    """Ordens 1 e 2 em 14/12/2026 (ganho 3.000,00: 450,00, vence 29/01/2027) e 3 e 4 em
    18/01/2027 (ganho 7.000,00: 1.050,00, vence no último dia útil de fevereiro de 2027)."""
    livro = _livro()
    janeiro = date(2027, 1, 18)
    datas = (DEZ, DEZ, janeiro, janeiro)
    r = _simular(livro, [_ordem(livro, n, "XXXX3", 2500, p, datas[n - 1]) for p, n in ESCADA],
                 date(2027, 1, 31))  # fmt: skip
    assert _mes(r, 2026, 12) == (D("450.00"), D("450.00"), date(2027, 1, 29))
    assert _mes(r, 2027, 1) == (D("1050.00"), D("1050.00"), date(2027, 2, 26))
    assert r.irrf_alienacao == D("5.50")  # 2,65 em dezembro e 2,85 em janeiro
