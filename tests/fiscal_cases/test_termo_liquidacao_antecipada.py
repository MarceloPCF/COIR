"""Termo liquidado antes do vencimento (parecer do especialista de 17/11/2026, casos TA-1 a TA-3).

- A alienação a termo se dá "no vencimento ou liquidação antecipada" (IN RFB 1.585/2015 art. 59
  §2º IV); na compra a termo, o custo é o preço a termo e a data de aquisição é a da liquidação
  (art. 62 I e §1º). Na liquidação antecipada o comprador paga o valor contratado: não há
  desconto de juros, e o custo continua sendo o preço a termo mais os custos da nota.
- A data que vale é a da liquidação (entrega das ações), não a do pedido: a da linha
  "Liquidação Termo" da movimentação da B3, que o serviço lê (`coir/servicos/termo.py`).
- Liquidação antecipada parcial: cada parte é uma liquidação própria, pelo preço a termo vezes a
  quantidade liquidada; os custos da nota se repartem na proporção (art. 56 §4º, por analogia).
  O serviço entrega ao motor uma operação por parte, como aqui.

Números conferidos à mão no parecer; sem IRRF; imposto truncado; DARF 6015."""

from datetime import date

from coir.domain.fiscal import Categoria, Grupo
from coir.engine.movimentos import termos_em_aberto
from tests.fiscal_cases.apoio import D, Livro, categoria

TERMO = "TERMO"


def test_ta_1_compra_a_termo_antecipada_entra_no_pm_na_data_da_liquidacao() -> None:
    """05/01/2026: 1.000 PETR4 a 30,00. 02/03: compra a termo de 1.000 a 28,00, prazo 90 dias
    (vencimento 01/06), custos 30,00; liquidada antes, em 14/04. 20/04: venda à vista de 1.000 a
    35,00, custos 20,00. Em 14/04: 2.000 ações por 58.030,00 (PM 29,015). Abril: 34.980,00 -
    29.015,00 = 5.965,00; IR 894,75 (vence 29/05/2026). Sobram 1.000 por 29.015,00. Com a data do
    vencimento, a venda de abril sairia só das ações antigas (ganho 4.980,00, IR 747,00)."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "PETR4", 1000, "30"))
    livro.nota(date(2026, 3, 2), livro.op("C", "PETR4", 1000, "28", mercado=TERMO,
                                          liquidacao=date(2026, 4, 14)), custos="30")  # fmt: skip
    livro.nota(date(2026, 4, 20), livro.op("V", "PETR4", 1000, "35"), custos="20")
    [p] = livro.carteira(date(2026, 4, 14))
    assert (p.quantidade, p.custo) == (D(2000), D(58030))
    abril = categoria(livro.apurar(date(2026, 4, 30)), 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.imposto) == (D(5965), D("894.75"))
    [p] = livro.carteira(date(2026, 4, 30))
    assert (p.quantidade, p.custo) == (D(1000), D(29015))


def test_ta_2_venda_a_termo_antecipada_vai_para_o_mes_e_o_limite_da_liquidacao() -> None:
    """1.000 VALE3 a 60,00 e 100 ITSA4 a 10,00. 02/03/2026: venda a termo de 1.000 VALE3 a
    66,00, prazo 60 dias (vencimento 04/05, 01/05 é feriado), custos 25,00; liquidada em 30/03.
    10/03: venda de 100 ITSA4 a 12,00. Março: soma do limite 66.000,00 + 1.200,00 = 67.200,00,
    a ITSA4 é tributada; resultado 5.975,00 + 200,00 = 6.175,00, IR 926,25 (vence 30/04/2026).
    Maio fica sem nada."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "VALE3", 1000, "60"),
               livro.op("C", "ITSA4", 100, "10"))  # fmt: skip
    livro.nota(date(2026, 3, 2), livro.op("V", "VALE3", 1000, "66", mercado=TERMO,
                                          liquidacao=date(2026, 3, 30)), custos="25")  # fmt: skip
    livro.nota(date(2026, 3, 10), livro.op("V", "ITSA4", 100, "12"))
    ap = livro.apurar(date(2026, 5, 31))
    marco, maio = (categoria(ap, 2026, m, Categoria.COMUM) for m in (3, 5))
    assert (marco.vendas_limite, marco.isento) == (D(67200), D(0))
    assert (marco.resultado, marco.imposto) == (D(6175), D("926.25"))
    assert marco.por_grupo[Grupo.TERMO] == D(5975)
    assert (maio.resultado, maio.vendas_limite) == (D(0), D(0))


def _parcial(*partes: tuple[int, date]) -> Livro:
    """04/05/2026: compra a termo de 1.000 BBAS3 a 25,00, prazo 60 dias (vencimento 03/07),
    custos 40,00, liquidada nas partes dadas (uma operação por parte, como o serviço monta)."""
    livro = Livro()
    livro.nota(date(2026, 5, 4), *(livro.op("C", "BBAS3", q, "25", mercado=TERMO, liquidacao=dia)
                                   for q, dia in partes), custos="40")  # fmt: skip
    return livro


def test_ta_3_antecipada_parcial_cada_parte_na_sua_data() -> None:
    """400 ações liquidadas em 16/06 (10.000,00 + 16,00 de custos) e vendidas à vista no mesmo
    dia a 27,00, custos 10,00: 10.790,00 - 10.016,00 = 774,00 na linha do termo, sem isenção
    mesmo com vendas de 10.800,00 (art. 59 §2º IV); IR 116,10 em junho (vence 31/07/2026). As
    outras 600 entram em 03/07 por 15.024,00."""
    livro = _parcial((400, date(2026, 6, 16)), (600, date(2026, 7, 3)))
    livro.nota(date(2026, 6, 16), livro.op("V", "BBAS3", 400, "27"), custos="10")
    junho = categoria(livro.apurar(date(2026, 6, 30)), 2026, 6, Categoria.COMUM)
    assert (junho.resultado, junho.isento, junho.imposto) == (D(774), D(0), D("116.10"))
    assert junho.por_grupo[Grupo.TERMO] == D(774)
    assert livro.carteira(date(2026, 7, 2)) == []
    [aberto] = termos_em_aberto(livro.notas, date(2026, 7, 2))
    assert (aberto.quantidade, aberto.custo, aberto.liquidacao) == (D(600), D(15024),
                                                                    date(2026, 7, 3))  # fmt: skip
    [p] = livro.carteira(date(2026, 7, 3))
    assert (p.quantidade, p.custo) == (D(600), D(15024))


def test_ta_3_variante_tres_partes_custos_pela_proporcao() -> None:
    """Três liquidações (333, 333 e 334 ações): custos de 13,32, 13,32 e 13,36, que fecham os
    40,00 da nota; ao fim, 1.000 ações por 25.040,00."""
    livro = _parcial((333, date(2026, 5, 18)), (333, date(2026, 6, 1)), (334, date(2026, 7, 3)))
    custos = [t.custo - t.quantidade * 25 for t in termos_em_aberto(livro.notas, date(2026, 5, 4))]
    assert custos == [D("13.32"), D("13.32"), D("13.36")]
    [p] = livro.carteira(date(2026, 6, 1))
    assert (p.quantidade, p.custo) == (D(666), D("16676.64"))
    [p] = livro.carteira(date(2026, 7, 3))
    assert (p.quantidade, p.custo) == (D(1000), D(25040))


def test_ta_3_variante_custo_nao_divisivel_fecha_com_a_nota() -> None:
    """Custos de 40,01 em duas partes de 500: o motor rateia os centavos pelo maior resto (20,01 e
    20,00, empate para a primeira), e a soma fecha com a nota: 1.000 ações por 25.040,01."""
    livro = Livro()
    livro.nota(date(2026, 5, 4),
               livro.op("C", "BBAS3", 500, "25", mercado=TERMO, liquidacao=date(2026, 6, 16)),
               livro.op("C", "BBAS3", 500, "25", mercado=TERMO, liquidacao=date(2026, 7, 3)),
               custos="40.01")  # fmt: skip
    [p] = livro.carteira(date(2026, 6, 16))
    assert (p.quantidade, p.custo) == (D(500), D("12520.01"))
    [p] = livro.carteira(date(2026, 7, 3))
    assert (p.quantidade, p.custo) == (D(1000), D("25040.01"))
