"""Soma do limite de isenção de R$ 20 mil: toda venda de ação do mês entra, inclusive em day trade,
por exercício de opção e na liquidação do termo; a isenção alcança só o ganho das vendas comuns.
Números conferidos à mão, sem custos.

Fontes: Lei 11.033/2004 art. 3º I ("valor das alienações [...] para o conjunto de ações"); IN RFB
1.585/2015 art. 59 I ("total das alienações desse ativo") e §2º I e IV (a isenção não se aplica a
day trade, exercício e termo: tira o benefício, não manda deixar de computar); art. 67 §3º (só a
ação de PME tem regra de não computar no limite); art. 60 II b (prêmio do lançador na alienação);
art. 63 §3º II a (exercício sem retenção); art. 65 §§8º e 11 (day trade: 20% e IRRF de 1%); CTN
art. 111 II (isenção interpretada literalmente). Critério do COIR (leitura literal), parâmetro
day_trade_na_soma_limite = true.
"""

from datetime import date
from decimal import Decimal

from coir.domain.fiscal import Categoria, OperacaoMotor
from tests.fiscal_cases.apoio import Livro, categoria, mes, parametros

D = Decimal


def _d(o: OperacaoMotor) -> OperacaoMotor:
    return o.model_copy(update={"marcada_dt": True})


def _day_trade(venda_comum: str) -> Livro:
    """Fevereiro: compra 1.000 AAAA3 a 10,00. 10/03: venda comum de 1.000 AAAA3 ao preço dado e day
    trade (marcado D) de 400 BBBB3, compra a 19,50 (7.800,00) e venda a 20,00 (8.000,00): ganho
    200,00, IRRF de 1% = 2,00 impresso na nota."""
    livro = Livro()
    livro.nota(date(2026, 2, 2), livro.op("C", "AAAA3", 1000, "10"))
    livro.nota(date(2026, 3, 10), livro.op("V", "AAAA3", 1000, venda_comum),
               _d(livro.op("C", "BBBB3", 400, "19.50")), _d(livro.op("V", "BBBB3", 400, "20")),
               irrf_dt="2.00")  # fmt: skip
    return livro


def test_venda_em_day_trade_entra_na_soma_e_o_ganho_comum_perde_a_isencao() -> None:
    """Caso A. Venda comum 15.000,00 (ganho 5.000,00; IRRF de 0,005% = 0,75, dispensado) + day trade
    8.000,00 = 23.000,00 > 20.000,00: isento 0. Imposto comum 15% x 5.000,00 = 750,00; day trade
    20% x 200,00 = 40,00; total 790,00 - IRRF 2,00 = DARF 788,00, vencimento 30/04/2026. Com o
    parâmetro desligado (contraprova): soma 15.000,00, ganho comum isento, DARF 38,00."""
    livro = _day_trade("15")
    ap = livro.apurar(date(2026, 3, 31))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.vendas_limite, comum.isento, comum.imposto) == (
        D("15000.00"), D("23000.00"), 0, D("750.00"))  # fmt: skip
    day_trade = categoria(ap, 2026, 3, Categoria.DAY_TRADE)
    assert (day_trade.resultado, day_trade.imposto) == (D("200.00"), D("40.00"))
    m = mes(ap, 2026, 3)
    assert (m.imposto, m.darf, m.vencimento) == (D("790.00"), D("788.00"), date(2026, 4, 30))
    [aviso] = ap.avisos
    assert aviso.startswith("03/2026: as vendas de ações em day trade (R$ 8.000,00) entram")
    assert livro.carteira(date(2026, 3, 31)) == []

    ap = livro.apurar(date(2026, 3, 31), parametros(day_trade_na_soma_limite="false"))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_limite, comum.isento, comum.imposto) == (D("15000.00"), D("5000.00"), 0)
    assert mes(ap, 2026, 3).darf == D("38.00")


def test_soma_igual_ao_limite_continua_isenta() -> None:
    """Caso A2. Venda comum 12.000,00 (ganho 2.000,00) + day trade 8.000,00 = 20.000,00, "igual ou
    inferior" ao limite: ganho comum isento; DARF 40,00 - 2,00 = 38,00."""
    ap = _day_trade("12").apurar(date(2026, 3, 31))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_limite, comum.isento, comum.imposto) == (D("20000.00"), D("2000.00"), 0)
    assert mes(ap, 2026, 3).darf == D("38.00")
    assert not ap.avisos


def test_um_centavo_acima_do_limite_tributa_o_ganho_comum() -> None:
    """Caso A3. Venda comum a 12,0001 = 12.000,10 (ganho 2.000,10) + day trade 8.000,00 =
    20.000,10: 15% x 2.000,10 = 300,015 -> 300,01 (truncado); DARF 300,01 + 40,00 - 2,00 =
    338,01."""
    ap = _day_trade("12.0001").apurar(date(2026, 3, 31))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_limite, comum.isento, comum.imposto) == (D("20000.10"), 0, D("300.01"))
    assert mes(ap, 2026, 3).darf == D("338.01")


def test_venda_por_exercicio_entra_na_soma_e_o_ganho_nunca_e_isento() -> None:
    """Caso B. Posição: 1.000 BBBB3 a 25,00 e 1.000 AAAA3 a 10,00. 02/06: lança 400 BBBBF300 a 0,50
    (prêmio 200,00). 15/06: exercício, venda de 400 BBBB3 a 30,00 = 12.000,00; alienação 12.000,00
    + 200,00 = 12.200,00; custo 10.000,00; ganho 2.200,00 (nunca isento). 20/06: venda comum de
    1.000 AAAA3 a 15,00 = 15.000,00 (ganho 5.000,00). Soma do limite: 15.000,00 + 12.000,00 (o
    prêmio não é venda de ação) = 27.000,00 > 20.000,00. Imposto 15% x 7.200,00 = 1.080,00; sem
    IRRF (exercício sem retenção; 0,75 da venda comum dispensado). DARF 1.080,00 em 31/07/2026.
    Fica 600 BBBB3 a 25,00 (15.000,00)."""
    livro = Livro()
    livro.nota(
        date(2026, 5, 4), livro.op("C", "BBBB3", 1000, "25"), livro.op("C", "AAAA3", 1000, "10")
    )
    livro.nota(date(2026, 6, 2), livro.op("V", "BBBBF300", 400, "0.50", mercado="CALL",
                                          vencimento=date(2026, 6, 15)))  # fmt: skip
    livro.nota(date(2026, 6, 15), livro.op("V", "BBBB3", 400, "30", mercado="EXERC CALL",
                                           opcao="BBBBF300"))  # fmt: skip
    livro.nota(date(2026, 6, 20), livro.op("V", "AAAA3", 1000, "15"))
    ap = livro.apurar(date(2026, 6, 30))
    comum = categoria(ap, 2026, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.vendas_limite, comum.resultado, comum.isento) == (
        D("27000.00"), D("27000.00"), D("7200.00"), 0)  # fmt: skip
    m = mes(ap, 2026, 6)
    assert (m.imposto, m.darf, m.vencimento) == (D("1080.00"), D("1080.00"), date(2026, 7, 31))
    [pos] = livro.carteira(date(2026, 6, 30))
    assert (pos.ticker, pos.quantidade, pos.custo) == ("BBBB3", 600, D("15000.00"))
