"""IRRF por ficha do programa IRPF com a DARF no saldo único (decisão do usuário de 16/10/2026:
separar o IRRF de FII só na declaração). Casos do especialista fiscal, conferidos à mão.

Fontes: Lei 11.033/2004 art. 2º §§1º, 2º, 5º e 7º; IN RFB 1.585/2015 arts. 63 §8º e 65 §§8º e 9º;
fichas "Renda variável - operações comuns/day trade" e "Operações - fundos de investimento
imobiliário" do programa IRPF. Custos zero para isolar o IRRF.
"""

from datetime import date
from decimal import Decimal

from hypothesis import given
from hypothesis import strategies as st

from coir.domain.fiscal import Categoria, Ficha, TipoAtivo
from tests.fiscal_cases.apoio import D, Livro, categoria, mes

FII = TipoAtivo.FII


def test_irrf_de_fii_usado_pela_ficha_de_comuns_fica_nela() -> None:
    """Março/2026: V 1.000 XPML11 a 100 (PM 100, resultado 0) e V 1.000 VALE3 a 60 (PM 50, ganho
    10.000; vendas 60.000 > 20.000: 15% = 1.500,00), IRRF 8,00 = 5,00 de FII (8 x 100/160) + 3,00.
    O saldo único deduz os 8,00 do imposto de comuns: DARF 1.492,00 (vence 30/04/2026). Na ficha
    de comuns, IR fonte no mês 8,00 e a pagar 1.492,00; a de FII fica sem IRRF.
    Abril: V 500 XPML11 a 110 (ganho 5.000; 20% = 1.000,00), IRRF 2,75: DARF 997,25."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "XPML11", 1500, "100", tipo=FII),
               livro.op("C", "VALE3", 1000, "50"))  # fmt: skip
    livro.nota(date(2026, 3, 10), livro.op("V", "XPML11", 1000, "100", tipo=FII),
               livro.op("V", "VALE3", 1000, "60"), irrf="8.00")  # fmt: skip
    livro.nota(date(2026, 4, 10), livro.op("V", "XPML11", 500, "110", tipo=FII), irrf="2.75")
    ap = livro.apurar(date(2026, 4, 30))
    marco, abril = mes(ap, 2026, 3), mes(ap, 2026, 4)
    assert (marco.darf, marco.vencimento) == (D("1492.00"), date(2026, 4, 30))
    comuns, fii = marco.ficha(Ficha.COMUNS), marco.ficha(Ficha.FII)
    assert (comuns.imposto, comuns.irrf_mes, comuns.a_pagar, comuns.pago) == (
        D("1500.00"), D("8.00"), D("1492.00"), D("1492.00"))  # fmt: skip
    assert (fii.irrf_mes, fii.irrf_a_compensar, fii.a_pagar) == (0, 0, 0)
    fii = abril.ficha(Ficha.FII)
    assert (abril.darf, fii.imposto, fii.irrf_mes, fii.a_pagar, fii.pago) == (
        D("997.25"), D("1000.00"), D("2.75"), D("997.25"), D("997.25"))  # fmt: skip


def test_nota_mista_reparte_o_irrf_pelas_vendas_comuns() -> None:
    """12/05/2026, uma nota: V 300 HGLG11 a 165,30 = 49.590 (PM 160, ganho 1.590), V 1.000 ITUB4
    a 33,47 = 33.470 (PM 30, ganho 3.470), V 2.000 PETRF350 a 0,85 = 1.700 (PM 0,60, ganho 500) e
    day trade BBAS3 500 a 25,00/25,40 (ganho 200, IRRF 1% 2,00). IRRF 0,005% 4,23 (base 84.760,
    sem o day trade): FII 4,23 x 49.590/84.760 = 2,47; comuns 1,76.
    Comuns: 15% x 3.970 = 595,50 + 20% x 200 = 40,00 = 635,50 - 2,00 - 1,76 = 631,74.
    FII: 20% x 1.590 = 318,00 - 2,47 = 315,53. DARF 947,27 (igual ao saldo único)."""
    livro = Livro()
    v = date(2026, 6, 19)
    livro.nota(date(2026, 4, 1), livro.op("C", "HGLG11", 300, "160", tipo=FII),
               livro.op("C", "ITUB4", 1000, "30"),
               livro.op("C", "PETRF350", 2000, "0.60", mercado="CALL", vencimento=v))  # fmt: skip
    livro.nota(date(2026, 5, 12), livro.op("V", "HGLG11", 300, "165.30", tipo=FII),
               livro.op("V", "ITUB4", 1000, "33.47"),
               livro.op("V", "PETRF350", 2000, "0.85", mercado="CALL", vencimento=v),
               livro.op("C", "BBAS3", 500, "25"), livro.op("V", "BBAS3", 500, "25.40"),
               irrf="4.23", irrf_dt="2.00")  # fmt: skip
    m = mes(ap := livro.apurar(date(2026, 5, 31)), 2026, 5)
    assert categoria(ap, 2026, 5, Categoria.FII).imposto == D("318.00")
    comuns, fii = m.ficha(Ficha.COMUNS), m.ficha(Ficha.FII)
    assert (comuns.imposto, comuns.irrf_day_trade_mes, comuns.irrf_mes, comuns.a_pagar) == (
        D("635.50"), D("2.00"), D("1.76"), D("631.74"))  # fmt: skip
    assert (fii.irrf_mes, fii.a_pagar) == (D("2.47"), D("315.53"))
    assert m.darf == D("947.27") == comuns.a_pagar + fii.a_pagar


def test_irrf_de_fii_usado_em_dezembro_nao_vai_para_a_declaracao() -> None:
    """10/12/2026: V 250 KNRI11 a 160 = 40.000 (PM 168, prejuízo FII -2.000), V 300 WEGE3 a 50 =
    15.000 (ganho 1.200, isento: vendas <= 20.000) e day trade PETR4 1.000 a 30,00/30,50 (ganho
    500; 20% = 100,00; IRRF 1% 5,00). IRRF 0,005% 2,75: FII 2,00 (40/55) e comuns 0,75. O saldo
    único deduz 5,00 + 0,75 + 2,00: DARF 92,25 (vence 29/01/2027), nada para a declaração; os
    2,00 de origem FII ficam na ficha de comuns (sem uso duplo)."""
    livro = Livro()
    livro.nota(date(2026, 11, 3), livro.op("C", "KNRI11", 250, "168", tipo=FII),
               livro.op("C", "WEGE3", 300, "46"))  # fmt: skip
    livro.nota(date(2026, 12, 10), livro.op("V", "KNRI11", 250, "160", tipo=FII),
               livro.op("V", "WEGE3", 300, "50"), livro.op("C", "PETR4", 1000, "30"),
               livro.op("V", "PETR4", 1000, "30.50"), irrf="2.75", irrf_dt="5.00")  # fmt: skip
    ap = livro.apurar(date(2026, 12, 31))
    dez = mes(ap, 2026, 12)
    assert (dez.darf, dez.vencimento) == (D("92.25"), date(2027, 1, 29))
    comuns, fii = dez.ficha(Ficha.COMUNS), dez.ficha(Ficha.FII)
    assert (comuns.irrf_day_trade_mes, comuns.irrf_mes, comuns.a_pagar) == (
        D("5.00"), D("2.75"), D("92.25"))  # fmt: skip
    assert (fii.irrf_mes, fii.irrf_a_compensar) == (0, 0)
    [anual] = ap.irrf_anual
    assert anual.saldo == 0
    assert anual.saldo_comum_fichas == {Ficha.COMUNS: 0, Ficha.FII: 0}


def test_sobra_de_irrf_de_fii_fica_na_ficha_fii() -> None:
    """Dezembro/2026, uma nota: V 2.000 opções a 0,75 = 1.500 (PM 0,745, ganho 10; 15% = 1,50) e
    V 250 cotas de FII a 160 = 40.000 (resultado 0), IRRF 2,07: FII 2,07 x 40.000/41.500 = 1,995
    -> 2,00; comuns 0,07. O imposto de 1,50 consome 0,07 + 1,43 de origem FII: a ficha de comuns
    tem IR fonte 1,50 e a pagar 0; a de FII fica com 0,57 a compensar, que vai para a declaração
    (o saldo único é 0,57; separar só pela origem levaria 2,00, usando 1,43 duas vezes)."""
    livro = Livro()
    venc = date(2026, 12, 18)
    livro.nota(date(2026, 12, 1), livro.op("C", "PETRL350", 2000, "0.745", mercado="CALL",
                                           vencimento=venc),
               livro.op("C", "XPML11", 250, "160", tipo=FII))  # fmt: skip
    livro.nota(date(2026, 12, 10), livro.op("V", "PETRL350", 2000, "0.75", mercado="CALL",
                                            vencimento=venc),
               livro.op("V", "XPML11", 250, "160", tipo=FII), irrf="2.07")  # fmt: skip
    ap = livro.apurar(date(2026, 12, 31))
    dez = mes(ap, 2026, 12)
    comuns, fii = dez.ficha(Ficha.COMUNS), dez.ficha(Ficha.FII)
    assert (dez.darf, dez.irrf_saldo) == (0, D("0.57"))
    assert (comuns.imposto, comuns.irrf_mes, comuns.a_pagar) == (D("1.50"), D("1.50"), 0)
    assert (fii.irrf_mes, fii.irrf_a_compensar) == (D("0.57"), D("0.57"))
    [anual] = ap.irrf_anual
    assert (anual.saldo_comum, anual.saldo_comum_fichas) == (
        D("0.57"), {Ficha.COMUNS: 0, Ficha.FII: D("0.57")})  # fmt: skip


def test_irrf_de_day_trade_nao_abate_fii_antes_do_0005() -> None:
    """IN RFB 1.585/2015 art. 65 §8º I e ficha FII sem linha de day trade: dezembro/2026 com
    imposto só de FII (V 100 XPML11 a 100,50, PM 100: ganho 50, 20% = 10,00), IRRF 0,005% de 20,00
    na venda de FII e IRRF de day trade de 5,00. A ficha FII consome 10,00 do 0,005% dela; o day
    trade fica para restituição (5,00) e 10,00 de 0,005% para a declaração."""
    livro = Livro()
    livro.nota(date(2026, 11, 3), livro.op("C", "XPML11", 100, "100", tipo=FII))
    livro.nota(date(2026, 12, 1), livro.op("C", "AAAA3", 10, "10"), irrf_dt="5")
    livro.nota(date(2026, 12, 10), livro.op("V", "XPML11", 100, "100.50", tipo=FII), irrf="20")
    ap = livro.apurar(date(2026, 12, 31))
    [anual] = ap.irrf_anual
    assert (anual.saldo_day_trade, anual.saldo_comum) == (D(5), D(10))
    assert anual.saldo_comum_fichas == {Ficha.COMUNS: 0, Ficha.FII: D(10)}
    assert mes(ap, 2026, 12).ficha(Ficha.COMUNS).irrf_day_trade_a_compensar == D(5)


def test_darf_acumulada_marca_os_meses_que_ela_pagou() -> None:
    """Lei 9.430/1996 art. 68 §1º: fevereiro/2026 com imposto 6,00 (ganho 40 em ações; vendas
    acima de 20.000) acumula; março soma 8,25 e paga 14,25. O "pago" de cada ficha é o a pagar do
    próprio mês quando a DARF que o cobriu já foi emitida (fevereiro e março)."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "AAAA3", 2000, "20"), livro.op("C", "BBBB3", 2000,
                                                                              "20"))  # fmt: skip
    livro.nota(date(2026, 2, 10), livro.op("V", "AAAA3", 2000, "20.02"))
    livro.nota(date(2026, 3, 10), livro.op("V", "BBBB3", 2000, "20.0275"))
    ap = livro.apurar(date(2026, 3, 31))
    fev, mar = mes(ap, 2026, 2), mes(ap, 2026, 3)
    assert (fev.darf, fev.darf_acumulado, mar.darf) == (0, D("6.00"), D("14.25"))
    assert fev.ficha(Ficha.COMUNS).pago == D("6.00")
    assert mar.ficha(Ficha.COMUNS).pago == D("8.25")
    parcial = livro.apurar(date(2026, 2, 28))
    assert mes(parcial, 2026, 2).ficha(Ficha.COMUNS).pago == 0


def test_base_da_reparticao_usa_o_bruto_e_ignora_venda_por_exercicio() -> None:
    """Lei 11.033/2004 art. 2º §2º: venda por exercício fora da base do 0,005%. Agosto/2026, uma
    nota: V 1.000 AAAA3 a 30 por exercício de call lançada (prêmio 1.000), V 100 XPML11 a 100
    (bruto 10.000; PM 90: ganho 1.000, 20% = 200,00) e V 1.000 BBBB3 a 10 (bruto 10.000, PM 10),
    custos 100 rateados. IRRF 1,00: a base é 10.000 de FII + 10.000 de BBBB3 (o exercício e os
    custos ficam fora) -> FII 0,50 e comuns 0,50."""
    livro = Livro()
    livro.nota(date(2026, 7, 1), livro.op("C", "AAAA3", 1000, "25"),
               livro.op("C", "XPML11", 100, "90", tipo=FII), livro.op("C", "BBBB3", 1000, "10"),
               livro.op("V", "AAAAB30", 1000, "1", mercado="CALL",
                        vencimento=date(2026, 8, 21)))  # fmt: skip
    livro.nota(date(2026, 8, 21), livro.op("V", "AAAA3", 1000, "30", mercado="EXERC CALL",
                                           opcao="AAAAB30"),
               livro.op("V", "XPML11", 100, "100", tipo=FII), livro.op("V", "BBBB3", 1000, "10"),
               custos="100", irrf="1.00")  # fmt: skip
    ap = livro.apurar(date(2026, 8, 31))
    ago = mes(ap, 2026, 8)
    fii = ago.ficha(Ficha.FII)
    assert fii.imposto == categoria(ap, 2026, 8, Categoria.FII).imposto
    retido_fii = fii.irrf_mes + ago.ficha(Ficha.COMUNS).irrf_mes
    assert retido_fii == D("1.00")
    # comuns tem imposto do exercício: consome os 0,50 dela; FII consome os 0,50 dela
    assert (fii.irrf_mes, ago.ficha(Ficha.COMUNS).irrf_mes) == (D("0.50"), D("0.50"))


def test_darf_de_janeiro_paga_o_acumulado_de_dezembro() -> None:
    """Lei 9.430/1996 art. 68 §1º: dezembro/2025 com imposto 6,00 acumula e entra na DARF de
    janeiro/2026 (6,00 + 8,25 = 14,25): o "pago" de dezembro é 6,00 no relatório de 2025."""
    livro = Livro()
    livro.nota(date(2025, 11, 5), livro.op("C", "AAAA3", 2000, "20"),
               livro.op("C", "BBBB3", 2000, "20"))  # fmt: skip
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 2000, "20.02"))
    livro.nota(date(2026, 1, 12), livro.op("V", "BBBB3", 2000, "20.0275"))
    ap = livro.apurar(date(2026, 1, 31))
    dez, jan = mes(ap, 2025, 12), mes(ap, 2026, 1)
    assert (dez.darf, jan.darf) == (0, D("14.25"))
    assert (dez.ficha(Ficha.COMUNS).pago, jan.ficha(Ficha.COMUNS).pago) == (D("6.00"), D("8.25"))


def test_day_trade_usado_no_imposto_de_fii_gera_aviso() -> None:
    """IN RFB 1.585/2015 art. 65 §8º I: o IRRF de day trade deduz o imposto do mês (saldo único da
    DARF), mas a ficha FII do programa não tem a linha: dezembro/2026 com imposto só de FII
    (V 100 XPML11 a 105, PM 100: ganho 500, 20% = 100,00) e IRRF só de day trade (5,00) -> DARF
    95,00 e aviso."""
    livro = Livro()
    livro.nota(date(2026, 11, 3), livro.op("C", "XPML11", 100, "100", tipo=FII))
    livro.nota(date(2026, 12, 1), livro.op("C", "AAAA3", 10, "10"), irrf_dt="5")
    livro.nota(date(2026, 12, 10), livro.op("V", "XPML11", 100, "105", tipo=FII))
    ap = livro.apurar(date(2026, 12, 31))
    fii = mes(ap, 2026, 12).ficha(Ficha.FII)
    assert (mes(ap, 2026, 12).darf, fii.irrf_day_trade_mes, fii.a_pagar) == (
        D(95), D(5), D(95))  # fmt: skip
    assert any("day trade deduzido do imposto de FII" in a for a in ap.avisos)


@given(
    st.lists(
        st.tuples(
            st.integers(1, 12),  # mês
            st.sampled_from(["AAAA3", "XPML11"]),
            st.integers(1, 30).map(lambda q: q * 100),
            st.decimals("8", "14", places=2),
            st.decimals("0", "9", places=2),  # IRRF 0,005%
            st.decimals("0", "9", places=2),  # IRRF de day trade
        ),
        min_size=1,
        max_size=12,
    )
)
def test_fichas_fecham_com_o_saldo_unico(
    vendas: list[tuple[int, str, int, Decimal, Decimal, Decimal]],
) -> None:
    """Invariantes da atribuição: em todo mês, a soma do a pagar das fichas é o imposto devido do
    saldo único, a soma dos saldos a compensar é o saldo único, e nenhum saldo fica negativo."""
    livro = Livro()
    livro.nota(date(2025, 12, 1), livro.op("C", "AAAA3", 100_000, "10"),
               livro.op("C", "XPML11", 100_000, "10", tipo=FII))  # fmt: skip
    for i, (m, ticker, q, preco, irrf, dt) in enumerate(vendas):
        tipo = FII if ticker == "XPML11" else None
        livro.nota(date(2026, m, 1 + i), livro.op("V", ticker, q, str(preco), tipo=tipo),
                   irrf=str(irrf), irrf_dt=str(dt))  # fmt: skip
    ap = livro.apurar(date(2026, 12, 31))
    for m in ap.meses:
        if m.ano != 2026:
            continue
        fs = m.fichas
        assert sum((f.a_pagar for f in fs), D(0)) == m.imposto_devido
        saldo = sum((f.irrf_a_compensar + f.irrf_day_trade_a_compensar for f in fs), D(0))
        assert saldo == m.irrf_saldo
        for f in fs:
            assert min(f.irrf_a_compensar, f.irrf_day_trade_a_compensar, f.a_pagar) >= 0
