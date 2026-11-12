"""Direito, recibo e bônus de subscrição (tipo DIREITO): 15% sobre o ganho líquido, 20% em day
trade, sem isenção, e a venda do direito fica fora da soma do limite de R$ 20 mil (SC Cosit
145/2021 e 166/2021, ementas oficiais; Lei 11.033/2004 art. 2º; IN RFB 1.585/2015 arts. 58, 59 e
65). O direito recebido sem nota de compra entra com custo zero (IN 1.585 art. 58 §7º, por
analogia: parecer do especialista de 09/11/2026, seguido pelo usuário)."""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo
from tests.fiscal_cases.apoio import D, Livro, categoria, parametros

DIR = TipoAtivo.DIREITO


def test_direito_recebido_vendido_tem_custo_zero_e_paga_15() -> None:
    """Caso do especialista: 100 direitos recebidos, vendidos a 0,50 com 0,10 de custos. Ganho
    49,90; imposto 15% = 7,485, truncado 7,48; abaixo de R$ 10, fica para o mês seguinte."""
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("V", "ZZZZ12", 100, "0.50", tipo=DIR), custos="0.10")
    ap = livro.apurar(date(2025, 3, 31))
    marco = categoria(ap, 2025, 3, Categoria.COMUM)
    assert (marco.resultado, marco.vendas_acoes, marco.isento) == (D("49.90"), 0, 0)
    assert marco.imposto == D("7.48")
    entrada = next(m for m in ap.movimentos if m.origem.startswith("direito:"))
    assert (entrada.cv, entrada.quantidade, entrada.valor) == ("C", D(100), 0)
    assert "custo zero" in entrada.regra
    assert not [a for _, a in ap.avisos_titular if "maior que a posição" in a]


def test_venda_de_direito_nao_entra_no_limite_das_acoes() -> None:
    """SC 145/2021: vendas de ações 19.000 (isentas) + direitos 5.000 fora da soma; o ganho do
    direito (5.000, custo zero) paga 15% = 750,00."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "PETR4", 1000, "10"))
    livro.nota(date(2025, 2, 10), livro.op("V", "PETR4", 1000, "19"),
               livro.op("V", "PETR12", 1000, "5", tipo=DIR))  # fmt: skip
    fevereiro = categoria(livro.apurar(date(2025, 2, 28)), 2025, 2, Categoria.COMUM)
    assert fevereiro.vendas_acoes == D("19000")
    assert (fevereiro.isento, fevereiro.resultado, fevereiro.imposto) == (
        D("9000"), D("5000"), D("750.00"))  # fmt: skip


def test_direito_comprado_usa_o_custo_pago_e_so_o_que_falta_tem_custo_zero() -> None:
    """Comprou 100 a 1,00 e vendeu 300 a 2,00 (200 recebidos): custo 100 + 0, ganho 500."""
    livro = Livro()
    livro.nota(date(2025, 4, 1), livro.op("C", "ZZZZ1", 100, "1", tipo=DIR))
    livro.nota(date(2025, 4, 8), livro.op("V", "ZZZZ1", 300, "2", tipo=DIR))
    ap = livro.apurar(date(2025, 4, 30))
    assert categoria(ap, 2025, 4, Categoria.COMUM).resultado == D(500)
    [entrada] = [m for m in ap.movimentos if m.origem.startswith("direito:")]
    assert entrada.quantidade == D(200)


def test_direito_de_fii_e_comum_a_15_nao_fii_a_20() -> None:
    """SC 166/2021: direito de cota de FII também é 15%, na ficha de operações comuns."""
    livro = Livro()
    livro.nota(date(2025, 5, 5), livro.op("V", "HGBS12", 1000, "1", tipo=DIR))
    ap = livro.apurar(date(2025, 5, 31))
    assert categoria(ap, 2025, 5, Categoria.COMUM).imposto == D("150.00")
    assert categoria(ap, 2025, 5, Categoria.FII).resultado == 0


def test_direito_comprado_e_vendido_no_dia_e_day_trade_a_20() -> None:
    """IN 1.585 art. 65 §1º I vale para qualquer ativo: 20% sobre 100."""
    livro = Livro()
    livro.nota(date(2025, 6, 2), livro.op("C", "ZZZZ1", 100, "1", tipo=DIR),
               livro.op("V", "ZZZZ1", 100, "2", tipo=DIR))  # fmt: skip
    ap = livro.apurar(date(2025, 6, 30))
    assert categoria(ap, 2025, 6, Categoria.DAY_TRADE).imposto == D("20.00")
    assert not [m for m in ap.movimentos if m.origem.startswith("direito:")]


def test_sem_o_parametro_a_venda_sem_compra_fica_a_descoberto() -> None:
    """`direito_recebido_custo_zero` é parâmetro com vigência: desligado, nada entra sozinho."""
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("V", "ZZZZ12", 100, "0.50", tipo=DIR))
    ap = livro.apurar(date(2025, 3, 31), parametros(direito_recebido_custo_zero="false"))
    assert not [m for m in ap.movimentos if m.origem.startswith("direito:")]
    assert categoria(ap, 2025, 3, Categoria.COMUM).resultado == 0
