"""Duas vendas no mesmo pregão que cruzam o zero: vale a ordem dos negócios.

Não há norma de "preço médio de venda": a IN RFB 1.585/2015 art. 58 trata do custo médio de
aquisição, e o único critério de ordem escrito é o do art. 65 §3º (primeira compra com a primeira
venda). A primeira venda zera a posição pelo custo médio; a seguinte abre a venda a descoberto
pelo próprio preço, com resultado na recompra e limite de isenção do mês da venda. A média das
vendas do dia que a corretora mostra é só exibição: o imposto total é o mesmo (parecer do
especialista de 16/11/2026; sem mudança de cálculo). Números conferidos à mão."""

from datetime import date

from coir.domain.fiscal import Categoria
from tests.fiscal_cases.apoio import D, Livro, categoria


def test_vendas_do_mesmo_pregao_que_cruzam_o_zero_seguem_a_ordem_dos_negocios() -> None:
    """3.000 ações a 7,50 (22.500,00). Em 10/03/2026: V 3.000 a 8,10 e depois V 3.000 a 8,20.
    Em 15/04/2026: recompra de 3.000 a 8,00.
    Março: 24.300,00 - 22.500,00 = 1.800,00; vendas de 48.900,00 > 20.000,00: imposto 270,00.
    Abril: 24.600,00 - 24.000,00 = 600,00, tributável pelo limite de março: imposto 90,00.
    Pela média do dia (8,15) seriam 1.950,00 e 450,00: o mesmo total de 2.400,00 e 360,00."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "ITSA4", 3000, "7.50"))
    livro.nota(date(2026, 3, 10), livro.op("V", "ITSA4", 3000, "8.10"),
               livro.op("V", "ITSA4", 3000, "8.20"))  # fmt: skip
    livro.nota(date(2026, 4, 15), livro.op("C", "ITSA4", 3000, "8.00"))
    ap = livro.apurar(date(2026, 4, 30))
    marco, abril = (categoria(ap, 2026, m, Categoria.COMUM) for m in (3, 4))
    assert (marco.resultado, marco.imposto) == (D(1800), D("270.00"))
    assert (abril.resultado, abril.imposto) == (D(600), D("90.00"))
    assert marco.resultado + abril.resultado == D(2400)
