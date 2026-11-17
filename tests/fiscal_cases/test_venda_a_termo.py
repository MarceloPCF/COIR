"""Venda a termo de ações, reconhecida na liquidação (item B7 do parecer do especialista de
13/11/2026).

- A alienação se dá na liquidação do contrato, no vencimento ou antecipada (IN RFB 1.585/2015 art.
  59 §2º IV), e não no pregão da venda.
- Ganho do vendedor descoberto: preço a termo - preço da compra à vista para a liquidação (art.
  62 II); do coberto: preço a termo - custo médio (art. 62 III; art. 58). O COIR usa o custo
  médio da posição na abertura do dia da liquidação: as compras feitas entre o contrato e a
  liquidação entram nele, e as duas leituras coincidem.
- Valor da alienação = preço a termo x quantidade - custos da nota da venda. Categoria comum, 15%,
  linha "Mercado a termo"; nunca isento (art. 59 §2º IV); o valor bruto entra na soma do limite
  de R$ 20 mil do mês da liquidação (leitura literal do art. 59 I, a mesma do exercício e do day
  trade).
- Compra à vista e venda a termo no mesmo pregão e corretora: possível operação conjugada (renda
  fixa, art. 47 I b): apurada como venda a termo, com aviso.

Casos T7-1 a T7-5 do parecer, conferidos à mão; sem IRRF, para simplificar. DARF 6015."""

from datetime import date

import pytest

from coir.domain.fiscal import Categoria, Grupo, Visao
from coir.erros import MotorError
from tests.fiscal_cases.apoio import D, Livro, categoria

TERMO = "TERMO"


def _coberta(liquidacao: date) -> Livro:
    """1.000 VALE3 a 60,00 (60.000,00). 02/03/2026: venda a termo de 1.000 a 66,00, prazo 30
    dias, custos 25,00."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "VALE3", 1000, "60"))
    livro.nota(date(2026, 3, 2), livro.op("V", "VALE3", 1000, "66", mercado=TERMO,
                                          liquidacao=liquidacao), custos="25")  # fmt: skip
    return livro


def test_t7_1_venda_a_termo_coberta_e_apurada_no_mes_da_liquidacao() -> None:
    """Liquidação em 01/04/2026. Abril: 66.000,00 - 25,00 - 60.000,00 = 5.975,00, linha do termo,
    não isento; soma do limite de abril: 66.000,00. Imposto 896,25 (vence 29/05/2026). Março
    fica sem resultado, e até a liquidação as ações continuam na carteira."""
    livro = _coberta(date(2026, 4, 1))
    [p] = livro.carteira(date(2026, 3, 31))
    assert (p.ticker, p.quantidade, p.custo) == ("VALE3", D(1000), D(60000))
    ap = livro.apurar(date(2026, 4, 30))
    marco, abril = (categoria(ap, 2026, m, Categoria.COMUM) for m in (3, 4))
    assert (marco.resultado, marco.vendas_limite) == (D(0), D(0))
    assert (abril.resultado, abril.isento, abril.imposto) == (D(5975), D(0), D("896.25"))
    assert abril.por_grupo[Grupo.TERMO] == D(5975) and abril.vendas_limite == D(66000)
    assert livro.carteira(date(2026, 4, 1)) == []


def test_t7_1_variante_com_venda_de_acoes_no_mes_da_liquidacao() -> None:
    """Mais uma venda de 100 ITSA4 a 12,00 (1.200,00, ganho 200,00) em abril: a soma passa de
    20.000,00 e os 200,00 são tributados. Imposto: 15% de 6.175,00 = 926,25."""
    livro = _coberta(date(2026, 4, 1))
    livro.nota(date(2026, 1, 6), livro.op("C", "ITSA4", 100, "10"))
    livro.nota(date(2026, 4, 10), livro.op("V", "ITSA4", 100, "12"))
    abril = categoria(livro.apurar(date(2026, 4, 30)), 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.isento, abril.imposto) == (D(6175), D(0), D("926.25"))


def test_t7_2_liquidacao_antecipada_leva_o_ganho_para_a_data_dela() -> None:
    """Liquidada antes do vencimento, em 16/03/2026: ganho de 5.975,00 em março, imposto 896,25
    (vence 30/04/2026)."""
    marco = categoria(_coberta(date(2026, 3, 16)).apurar(date(2026, 3, 31)), 2026, 3,
                      Categoria.COMUM)  # fmt: skip
    assert (marco.resultado, marco.imposto) == (D(5975), D("896.25"))


def test_t7_3_venda_a_termo_descoberta_com_compra_antes_da_liquidacao() -> None:
    """02/03/2026: venda a termo de 500 a 66,00, prazo 60 dias, custos 15,00; 01/05 é feriado,
    a liquidação fica em 04/05. 29/04: compra à vista de 500 a 61,00 + 10,00 = 30.510,00.
    Maio: 32.985,00 - 30.510,00 = 2.475,00; imposto 371,25 (vence 30/06/2026)."""
    livro = Livro()
    livro.nota(date(2026, 3, 2), livro.op("V", "VALE3", 500, "66", mercado=TERMO,
                                          liquidacao=date(2026, 5, 4)), custos="15")  # fmt: skip
    livro.nota(date(2026, 4, 29), livro.op("C", "VALE3", 500, "61"), custos="10")
    ap = livro.apurar(date(2026, 5, 31))
    abril, maio = (categoria(ap, 2026, m, Categoria.COMUM) for m in (4, 5))
    assert abril.resultado == D(0)
    assert (maio.resultado, maio.imposto) == (D(2475), D("371.25"))


def test_t7_4_compra_no_meio_do_contrato_entra_no_custo_medio() -> None:
    """Como o T7-1, mais uma compra em 20/03 de 500 a 64,00: posição 1.500 por 92.000,00. Na
    liquidação o custo das 1.000 é 61.333,33; ganho 65.975,00 - 61.333,33 = 4.641,67 (com o PM
    em 8 casas, 4.641,666...); imposto 696,25. Sobram 500 por 30.666,67."""
    livro = _coberta(date(2026, 4, 1))
    livro.nota(date(2026, 3, 20), livro.op("C", "VALE3", 500, "64"))
    ap = livro.apurar(date(2026, 4, 30))
    abril = categoria(ap, 2026, 4, Categoria.COMUM)
    assert abril.resultado.quantize(D("0.01")) == D("4641.67")
    assert abril.imposto == D("696.25")
    [p] = livro.carteira(date(2026, 4, 30))
    assert (p.quantidade, p.custo.quantize(D("0.01"))) == (D(500), D("30666.67"))


def test_t7_5_possivel_operacao_conjugada_e_apurada_com_aviso() -> None:
    """02/03/2026: compra à vista de 1.000 ITUB4 a 30,00 e venda a termo de 1.000 a 30,60, prazo
    30 dias, na mesma corretora. Apurada pelo art. 62 III: ganho de 600,00 em abril, imposto
    90,00, com aviso de possível operação conjugada (art. 47 I b)."""
    livro = Livro()
    livro.nota(date(2026, 3, 2), livro.op("C", "ITUB4", 1000, "30"),
               livro.op("V", "ITUB4", 1000, "30.60", mercado=TERMO,
                        liquidacao=date(2026, 4, 1)))  # fmt: skip
    ap = livro.apurar(date(2026, 4, 30))
    abril = categoria(ap, 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.imposto) == (D(600), D("90.00"))
    [aviso] = [a for a in ap.avisos if "operação conjugada" in a]
    assert "ITUB4" in aviso and "art. 47 I b" in aviso


def test_compra_e_venda_a_termo_liquidadas_no_mesmo_dia() -> None:
    """A compra a termo liquida antes da venda a termo do mesmo dia: as ações recebidas podem ser
    entregues. Compra a termo de 100 a 50,00 e venda a termo de 100 a 55,00, ambas liquidando em
    01/04/2026: ganho 500,00, não isento."""
    livro = Livro()
    livro.nota(date(2026, 3, 2), livro.op("C", "VALE3", 100, "50", mercado=TERMO,
                                          liquidacao=date(2026, 4, 1)))  # fmt: skip
    livro.nota(date(2026, 3, 3), livro.op("V", "VALE3", 100, "55", mercado=TERMO,
                                          liquidacao=date(2026, 4, 1)))  # fmt: skip
    abril = categoria(livro.apurar(date(2026, 4, 30)), 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.isento, abril.imposto) == (D(500), D(0), D("75.00"))


def test_venda_a_termo_de_lote_isento_de_pme_para_o_motor() -> None:
    """Sem norma para a isenção de PME (Lei 13.043/2014 art. 16) na entrega do termo e sem caso
    real: o motor para, sem palpite (regra 5)."""
    livro = Livro()
    livro.nota(date(2019, 1, 7), livro.op("C", "PMEE3", 100, "10", pme_ate=date(2023, 12, 31)))
    livro.nota(date(2019, 3, 1), livro.op("V", "PMEE3", 100, "12", mercado=TERMO,
                                          liquidacao=date(2019, 4, 1)))  # fmt: skip
    with pytest.raises(MotorError, match="lote isento de PME"):
        livro.movimentos(date(2019, 4, 30))


def test_compra_e_venda_a_termo_e_venda_a_vista_no_dia_da_liquidacao() -> None:
    """1.000 VALE3 a 60,00. Compra a termo de 100 a 50,00 e venda a termo de 100 a 55,00, ambas
    liquidando em 01/04/2026, e venda à vista de 100 a 70,00 nesse dia. A compra a termo entra
    primeiro (1.100 por 65.000,00). A venda a termo sai pelo custo médio (art. 62 III): 5.500,00
    - 5.909,09 = -409,09. A venda à vista do dia consome o lote do termo pelo preço a termo (art.
    62 I): 7.000,00 - 5.000,00 = 2.000,00. Nada é isento (art. 59 §2º IV): resultado 1.590,91,
    imposto 238,63. Ficam 900 ações por 54.090,91."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "VALE3", 1000, "60"))
    livro.nota(date(2026, 3, 2), livro.op("C", "VALE3", 100, "50", mercado=TERMO,
                                          liquidacao=date(2026, 4, 1)))  # fmt: skip
    livro.nota(date(2026, 3, 3), livro.op("V", "VALE3", 100, "55", mercado=TERMO,
                                          liquidacao=date(2026, 4, 1)))  # fmt: skip
    livro.nota(date(2026, 4, 1), livro.op("V", "VALE3", 100, "70"))
    abril = categoria(livro.apurar(date(2026, 4, 30)), 2026, 4, Categoria.COMUM)
    assert abril.resultado.quantize(D("0.01")) == D("1590.91")
    assert (abril.isento, abril.imposto) == (D(0), D("238.63"))
    [p] = livro.carteira(date(2026, 4, 1))
    assert (p.quantidade, p.custo.quantize(D("0.01"))) == (D(900), D("54090.91"))


def test_venda_a_termo_na_visao_por_conta() -> None:
    """A visão operacional (por conta) também liquida a venda a termo, na conta da nota."""
    livro = _coberta(date(2026, 4, 1))
    [termo] = [m for m in livro.movimentos(date(2026, 4, 30), Visao.CONTA)
               if m.origem.startswith("termo:")]  # fmt: skip
    assert (termo.data, termo.conta, termo.resultado) == (date(2026, 4, 1), 1, D(5975))
