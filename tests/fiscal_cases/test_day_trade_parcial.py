"""Day trade parcial: a corretora marca D a linha inteira, e só parte dela fecha no dia (parecer do
especialista fiscal de 07/11/2026; números conferidos à mão).

Fontes (IN RFB 1.585/2015):
- art. 65 §1º: day trade é a operação iniciada e encerrada no mesmo dia, com o mesmo ativo, na mesma
  instituição, "em que a quantidade negociada tenha sido liquidada, total ou parcialmente";
- art. 65 §2º: o estoque de antes do dia não conta;
- art. 65 §3º: pela ordem, o 1º negócio de compra com o 1º de venda, ou o inverso;
- art. 56 §§3º e 4º: dedução dos custos necessários, rateados proporcionalmente ao valor financeiro;
- art. 59 §2º I: o day trade fica fora da isenção; o excedente é operação comum.
"""

from datetime import date
from decimal import Decimal

from coir.domain.fiscal import Categoria, OperacaoMotor
from tests.fiscal_cases.apoio import Livro, categoria, mes, parametros

D = Decimal


def _d(o: OperacaoMotor) -> OperacaoMotor:
    return o.model_copy(update={"marcada_dt": True})


def test_compra_500_d_e_venda_200_d_dao_200_de_day_trade_e_300_de_compra_comum() -> None:
    """Sem posição antes. Compra 500 a 10,00 (5.000,00) e venda 200 a 10,50 (2.100,00), custos 7,10.

    Rateio pelo valor (7.100,00): compra 5,00 e venda 2,10.
    Compra com custos: 5.005,00; 200 -> 2.002,00 no day trade; 300 -> 3.003,00 na posição.
    Venda líquida: 2.097,90. Day trade: 2.097,90 - 2.002,00 = 95,90; imposto 20% = 19,18.
    IRRF de 1% impresso: 0,96. DARF: 19,18 - 0,96 = 18,22.
    Em abril, venda comum de 300 a 11,00: 3.300,00 - 3.003,00 = 297,00, isento (até R$ 20 mil)."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), _d(livro.op("C", "AAAA3", 500, "10")),
               _d(livro.op("V", "AAAA3", 200, "10.50")), custos="7.10", irrf_dt="0.96")  # fmt: skip
    livro.nota(date(2026, 4, 15), livro.op("V", "AAAA3", 300, "11"))
    ap = livro.apurar(date(2026, 4, 30))
    day_trade = categoria(ap, 2026, 3, Categoria.DAY_TRADE)
    assert (day_trade.resultado, day_trade.imposto) == (D("95.90"), D("19.18"))
    assert mes(ap, 2026, 3).darf == D("18.22")
    [posicao] = livro.carteira(date(2026, 3, 31))
    assert (posicao.quantidade, posicao.custo) == (D(300), D(3003))
    assert categoria(ap, 2026, 4, Categoria.COMUM).isento == D("297.00")
    assert not ap.avisos  # a marcação cobre o day trade do dia: nada a avisar


def test_venda_500_d_e_compra_200_d_dao_200_de_day_trade_e_300_de_venda_comum() -> None:
    """Posição de 1.000 a 9,00 (9.000,00). Venda 500 a 10,00 (5.000,00) e compra 200 a 9,50
    (1.900,00), custos 6,90.

    Rateio pelo valor (6.900,00): venda 5,00 e compra 1,90.
    Venda líquida: 4.995,00; 200 -> 1.998,00 no day trade; 300 -> 2.997,00 comuns (bruto 3.000,00).
    Compra com custos: 1.901,90. Day trade: 1.998,00 - 1.901,90 = 96,10; imposto 20% = 19,22.
    Venda comum: 2.997,00 - 300 x 9,00 = 297,00, isento: a soma do limite é o bruto comum, 3.000,00,
    mais a venda do day trade, 2.000,00 = 5.000,00 (Lei 11.033/2004 art. 3º I).
    O estoque não entra no day trade (§2º) e paga a venda comum: fica 700 a 9,00 = 6.300,00."""
    livro = Livro()
    livro.nota(date(2026, 2, 2), livro.op("C", "AAAA3", 1000, "9"))
    livro.nota(date(2026, 3, 10), _d(livro.op("V", "AAAA3", 500, "10")),
               _d(livro.op("C", "AAAA3", 200, "9.50")), custos="6.90")  # fmt: skip
    ap = livro.apurar(date(2026, 3, 31))
    day_trade = categoria(ap, 2026, 3, Categoria.DAY_TRADE)
    assert (day_trade.resultado, day_trade.imposto) == (D("96.10"), D("19.22"))
    assert mes(ap, 2026, 3).darf == D("19.22")
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.vendas_limite, comum.isento, comum.imposto) == (
        D("3000.00"),
        D("5000.00"),
        D("297.00"),
        0,
    )
    [posicao] = livro.carteira(date(2026, 3, 31))
    assert (posicao.quantidade, posicao.custo) == (D(700), D(6300))
    assert not ap.avisos


def test_linha_nao_marcada_antes_nao_entra_no_day_trade() -> None:
    """Compra 300 a 10,00 (sem D), compra 500 a 10,20 (D), venda 200 a 10,50 (D), sem custos.

    A linha marcada vem primeiro: day trade = 200 x (10,50 - 10,20) = 60,00. Pela ordem impressa
    seriam 100,00, com o negócio que a corretora não marcou. Sobram 300 a 10,00 (3.000,00) e 300 a
    10,20 (3.060,00): 600 com custo 6.060,00."""
    livro = Livro()
    livro.nota(
        date(2026, 3, 10),
        livro.op("C", "AAAA3", 300, "10"),
        _d(livro.op("C", "AAAA3", 500, "10.20")),
        _d(livro.op("V", "AAAA3", 200, "10.50")),
    )
    ap = livro.apurar(date(2026, 3, 31))
    assert categoria(ap, 2026, 3, Categoria.DAY_TRADE).resultado == D("60.00")
    [posicao] = livro.carteira(date(2026, 3, 31))
    assert (posicao.quantidade, posicao.custo) == (D(600), D(6060))
    assert not ap.avisos


def test_marcacao_menor_que_o_day_trade_do_dia_avisa_e_vale_a_ordem() -> None:
    """Compra 300 a 10,00 (D), venda 100 a 10,50 (D) e venda 200 a 10,60 (sem D), sem custos. O dia
    tem 300 de day trade e só 100 de venda marcada: a marcação não cobre o dia, vale a ordem
    impressa, com aviso. Day trade: 100 x 10,50 + 200 x 10,60 - 300 x 10,00 = 170,00."""
    livro = Livro()
    livro.nota(
        date(2026, 3, 10),
        _d(livro.op("C", "AAAA3", 300, "10")),
        _d(livro.op("V", "AAAA3", 100, "10.50")),
        livro.op("V", "AAAA3", 200, "10.60"),
    )
    ap = livro.apurar(date(2026, 3, 31))
    assert categoria(ap, 2026, 3, Categoria.DAY_TRADE).resultado == D("170.00")
    assert any("marcação D da nota não fecha" in a for a in ap.avisos)


def test_venda_em_day_trade_entra_na_soma_do_limite_de_isencao() -> None:
    """Lei 11.033/2004 art. 3º I e IN RFB 1.585/2015 art. 59 I ("total das alienações"); o §2º I
    só tira o day trade da isenção (parâmetro day_trade_na_soma_limite = true; critério do COIR).

    Março/2026: venda comum de 1.000 AAAA3 a 15,00 = 15.000,00 (PM 10,00, ganho 5.000,00) e day
    trade de 500 BBBB3 a 20,00/20,10 (venda 10.050,00, ganho 50,00, 20% = 10,00), sem custos.
    Soma do limite: 25.050,00 > 20.000,00, o ganho comum paga 15% = 750,00 e a DARF é 760,00, com
    aviso. Com o parâmetro desligado, a soma é 15.000,00, o ganho comum é isento e a DARF é
    10,00."""
    livro = Livro()
    livro.nota(date(2026, 2, 2), livro.op("C", "AAAA3", 1000, "10"))
    compra, venda = livro.op("C", "BBBB3", 500, "20"), livro.op("V", "BBBB3", 500, "20.10")
    livro.nota(date(2026, 3, 10), livro.op("V", "AAAA3", 1000, "15"), _d(compra), _d(venda))
    ap = livro.apurar(date(2026, 3, 31))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.vendas_limite, comum.isento, comum.imposto) == (
        D("15000.00"),
        D("25050.00"),
        0,
        D("750.00"),
    )
    assert mes(ap, 2026, 3).darf == D("760.00")
    [aviso] = ap.avisos
    assert aviso.startswith("03/2026: as vendas de ações em day trade (R$ 10.050,00) entram")
    ap = livro.apurar(date(2026, 3, 31), parametros(day_trade_na_soma_limite="false"))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.vendas_limite, comum.isento, comum.imposto) == (
        D("15000.00"),
        D("15000.00"),
        D("5000.00"),
        0,
    )
    assert not ap.avisos
    assert mes(ap, 2026, 3).darf == D("10.00")
