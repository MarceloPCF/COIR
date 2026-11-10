"""PETZ3 -> AUAU3 com R$ 0,71 por ação em dinheiro (parecer do especialista fiscal de 07/11/2026,
leitura B; números conferidos à mão).

Fatos (FAQ do RI Petz Cobasi): relação 1:1; último pregão de PETZ3 em 02/01/2026; AUAU3 desde
05/01/2026; cada ação valia R$ 4,39 no fechamento: uma ON de R$ 3,68 e uma PN resgatável de R$ 0,71,
paga em 23/01/2026. Valor exato pago: R$ 0,7109 por ação (imprensa, sobre o resultado do 4T25).

Regra:
- IN RFB 1.585/2015 art. 58 §6º: o custo das ações originais vai para as recebidas, e a ON e a PN
  foram recebidas. Repartição pelo valor atribuído pela companhia: 0,71 / 4,39 = 16,17312073%.
- O resgate da PN é alienação: dinheiro menos o custo da PN (Lei 7.713/1988 art. 3º §3º; SC Cosit
  37/2020, CETIP/BM&FBOVESPA, com a mesma estrutura de ON + PN resgatável). Fora de bolsa, vai para
  o GCAP e não entra na apuração mensal; a perda não compensa ganhos da bolsa.
- Isenção do resgate: até R$ 35 mil alienados no mês (Lei 9.250/1995 art. 22 II).
- A incorporação em si segue o parâmetro incorporacao_acoes_tratamento (custo mantido); a RFB
  diverge (SC Cosit 224/2014), e a tela de Eventos avisa.

O motor recebe só a repartição do custo, na data ex: REDUCAO_CAPITAL (V) pelo custo da PN, sem
resultado. O valor do evento é o custo da PN, e não o dinheiro recebido.
"""

from datetime import date
from decimal import ROUND_HALF_UP, Decimal

from coir import referencias
from coir.domain.fiscal import Categoria, NaturezaEvento, TipoEvento
from tests.fiscal_cases.apoio import Livro, categoria

D = Decimal
PARTE_DA_PN = D("0.1617312073")  # 0,71 / 4,39
DINHEIRO_POR_ACAO = D("0.7109")
DATA_EX = date(2026, 1, 5)


def _centavos(valor: Decimal) -> Decimal:
    return valor.quantize(D("0.01"), ROUND_HALF_UP)


def _custo_da_pn(custo: Decimal) -> Decimal:
    return _centavos(custo * PARTE_DA_PN)


def _com_evento(quantidade: int, preco: str) -> Livro:
    """Compra em 2025, sem custos; a incorporação 1:1 e a repartição do custo na data ex."""
    livro = Livro()
    custo = D(quantidade) * D(preco)
    livro.nota(date(2025, 6, 2), livro.op("C", "PETZ3", quantidade, preco))
    livro.evento(DATA_EX, "PETZ3", TipoEvento.INCORPORACAO, quantidade=str(quantidade),
                 para="AUAU3", quantidade_destino=str(quantidade),
                 natureza=NaturezaEvento.INCORPORACAO_ACOES)  # fmt: skip
    livro.evento(DATA_EX, "AUAU3", TipoEvento.REDUCAO_CAPITAL, "V", None, str(_custo_da_pn(custo)))
    return livro


def test_custo_acima_de_4_39_reparte_o_custo_e_o_resgate_da_perda_fora_da_bolsa() -> None:
    """100 ações a 5,00: custo 500,00. PN: 500,00 x 16,17312073% = 80,87. AUAU3: 419,13.
    Resgate: 100 x 0,7109 = 71,09; 71,09 - 80,87 = -9,78 no GCAP, que não compensa nada.
    Venda em fevereiro de 100 a 5,00: 500,00 - 419,13 = 80,87, isento (até R$ 20 mil).
    A leitura A daria custo de 428,91, e a C, resultado de -61,00 e custo de 368,00."""
    assert _custo_da_pn(D(500)) == D("80.87")
    assert _centavos(100 * DINHEIRO_POR_ACAO) - _custo_da_pn(D(500)) == D("-9.78")
    livro = _com_evento(100, "5")
    [auau] = livro.carteira(DATA_EX)
    assert (auau.ticker, auau.quantidade, auau.custo) == ("AUAU3", D(100), D("419.13"))
    assert all(m.resultado == 0 for m in livro.movimentos(DATA_EX))  # o GCAP fica fora da apuração
    livro.nota(date(2026, 2, 2), livro.op("V", "AUAU3", 100, "5"))
    comum = categoria(livro.apurar(date(2026, 2, 28)), 2026, 2, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(500), D("80.87"))


def test_custo_abaixo_de_4_39_o_resgate_da_ganho_isento_ate_35_mil() -> None:
    """100 ações a 2,00: custo 200,00. PN: 32,35. AUAU3: 167,65.
    Resgate: 71,09 - 32,35 = 38,74 de ganho, isento: R$ 71,09 alienados no mês."""
    assert _centavos(100 * DINHEIRO_POR_ACAO) - _custo_da_pn(D(200)) == D("38.74")
    [auau] = _com_evento(100, "2").carteira(DATA_EX)
    assert (auau.quantidade, auau.custo) == (D(100), D("167.65"))


def test_resgate_acima_de_35_mil_tem_imposto_no_gcap() -> None:
    """100.000 ações a 2,00: custo 200.000,00. PN: 32.346,24. AUAU3: 167.653,76.
    Resgate: 71.090,00 (acima de R$ 35 mil): ganho de 38.743,76; imposto de 15% = 5.811,56, por
    DARF até o último dia útil do mês seguinte ao do recebimento (27/02/2026)."""
    custo_pn = _custo_da_pn(D(200000))
    ganho = 100000 * DINHEIRO_POR_ACAO - custo_pn
    assert (custo_pn, ganho) == (D("32346.24"), D("38743.76"))
    assert _centavos(ganho * D("0.15")) == D("5811.56")
    [auau] = _com_evento(100000, "2").carteira(DATA_EX)
    assert auau.custo == D("167653.76")


def test_venda_na_data_ex_antes_do_pagamento_ja_usa_o_custo_repartido() -> None:
    """A ON e a PN separam-se no fechamento, e o dinheiro vai para quem tinha a ação em 02/01. A
    venda de 100 a 3,68 em 05/01: 368,00 - 419,13 = -51,13."""
    livro = _com_evento(100, "5")
    livro.nota(DATA_EX, livro.op("V", "AUAU3", 100, "3.68"))
    venda = livro.movimentos(date(2026, 1, 31))[-1]
    assert venda.resultado == D("-51.13")


def test_troca_de_codigo_na_tabela() -> None:
    refs = referencias.carregar()
    assert refs.ticker_na_data("PETZ3", date(2026, 1, 2)) == "AUAU3"
    assert refs.ticker_na_data("AUAU3", DATA_EX) == "AUAU3"
