"""Compra em oferta pública (IPO e follow-on), que não aparece em nota de corretagem.

IN RFB 1.585/2015 art. 58 caput: o custo é o valor pago, com preço médio ponderado; art. 56 §5º:
ganho apurado na alienação. Não há alienação na oferta: sem resultado e sem IRRF. Sem norma sobre a
data: usa-se a liquidação (parecer do especialista de 08/11/2026).
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, NaturezaEvento, TipoEvento
from coir.erros import MotorError
from tests.fiscal_cases.apoio import Livro, categoria, mes

D = Decimal


def test_follow_on_entra_no_preco_medio_e_a_venda_usa_o_novo_pm() -> None:
    """500 a 18,00 (9.000,00) + oferta de 1.000 a 20,00 (20.000,00) = 1.500 com custo de 29.000,00,
    PM 19,3333... Venda de 1.200 a 25,00 (30.000,00, acima de R$ 20 mil), sem custos: baixa de
    23.200,00, resultado de 6.800,00 e imposto de 1.020,00; restam 300 com custo de 5.800,00."""
    livro = Livro()
    livro.nota(date(2025, 8, 4), livro.op("C", "AAAA3", 500, "18"))
    livro.evento(date(2025, 9, 15), "AAAA3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    livro.nota(date(2025, 10, 8), livro.op("V", "AAAA3", 1200, "25"))
    ap = livro.apurar(date(2025, 10, 31))
    assert all(m.resultado == 0 for m in ap.movimentos if m.data == date(2025, 9, 15))
    outubro = categoria(ap, 2025, 10, Categoria.COMUM)
    assert (outubro.resultado, outubro.imposto) == (D("6800.00"), D("1020.00"))
    ultimo = ap.movimentos[-1]
    assert (ultimo.quantidade_apos, ultimo.custo_apos.quantize(D("0.01"))) == (D(300), D("5800.00"))


def test_ipo_abre_posicao_sem_nota_anterior() -> None:
    livro = Livro()
    livro.evento(date(2025, 3, 3), "BBBB3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    [mov] = livro.movimentos(date(2025, 3, 31))
    assert (mov.quantidade_apos, mov.pm_apos) == (D(1000), D(20))


def test_venda_no_dia_da_liquidacao_usa_o_custo_da_oferta() -> None:
    """Os eventos do dia vêm antes das operações: a oferta já está na posição. 1.000 a 26,00 =
    26.000,00; custo de 20.000,00; ganho de 6.000,00 e imposto de 900,00."""
    livro = Livro()
    livro.evento(date(2026, 6, 30), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    livro.nota(date(2026, 6, 30), livro.op("V", "CCCC3", 1000, "26"))
    junho = categoria(livro.apurar(date(2026, 6, 30)), 2026, 6, Categoria.COMUM)
    assert (junho.resultado, junho.imposto) == (D("6000.00"), D("900.00"))


def test_venda_antes_da_liquidacao_tem_a_compra_antes_dela() -> None:
    """Negociação em 30/06/2026 e liquidação em 01/07/2026, oferta de 1.000 a 20,00. Venda em 30/06
    de 1.000 a 26,00, custos 6,00: resultado 5.994,00 em junho, imposto 899,10, IRRF 1,30, DARF
    897,80, vencimento 31/07/2026; a posição termina zerada (decisão do usuário de 08/11/2026)."""
    for ate in (date(2026, 6, 30), date(2026, 7, 31)):  # a mesma resposta antes da liquidação
        livro = Livro()
        livro.nota(date(2026, 6, 30), livro.op("V", "CCCC3", 1000, "26"), custos="6", irrf="1.30")
        livro.evento(date(2026, 7, 1), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
        ap = livro.apurar(ate)
        junho = categoria(ap, 2026, 6, Categoria.COMUM)
        assert (junho.resultado, junho.imposto) == (D("5994.00"), D("899.10"))
        assert (mes(ap, 2026, 6).darf, mes(ap, 2026, 6).vencimento) == (
            D("897.80"),
            date(2026, 7, 31),
        )
        assert livro.carteira(ate) == []
        assert any("antes da liquidação" in a for a in ap.avisos)


def test_venda_de_acoes_antigas_antes_da_oferta_usa_o_pm_antigo() -> None:
    """500 a 18,00; venda de 200 a 25,00 três dias antes da liquidação do follow-on: a posição não
    fica vendida, então a compra não é antecipada. Ganho de 200 x 7,00 = 1.400,00 (isento: 5.000,00
    de venda), sem o custo da oferta."""
    livro = Livro()
    livro.nota(date(2025, 8, 4), livro.op("C", "AAAA3", 500, "18"))
    livro.nota(date(2025, 9, 12), livro.op("V", "AAAA3", 200, "25"))
    livro.evento(date(2025, 9, 15), "AAAA3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    setembro = categoria(livro.apurar(date(2025, 9, 30)), 2025, 9, Categoria.COMUM)
    assert setembro.isento == D("1400.00")


def test_posicao_vendida_ha_mais_tempo_falha_alto() -> None:
    """Fora da janela da liquidação não é venda das ações da oferta: o evento zeraria a posição e o
    ganho sumiria sem resultado (regra 5: falhar alto)."""
    livro = Livro()
    livro.nota(date(2026, 5, 4), livro.op("V", "CCCC3", 1000, "26"))
    livro.evento(date(2026, 7, 1), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    with pytest.raises(MotorError, match="há venda antes desta data"):
        livro.movimentos(date(2026, 7, 31))


def test_subscricao_com_posicao_vendida_falha_alto() -> None:
    """Sem a trava, a subscrição zerava a posição vendida e o ganho da venda sumia sem resultado."""
    livro = Livro()
    livro.nota(date(2026, 6, 1), livro.op("V", "CCCC3", 1000, "26"))
    livro.evento(date(2026, 7, 1), "CCCC3", TipoEvento.SUBSCRICAO, "C", "1000", "20000")
    with pytest.raises(MotorError, match="há venda antes desta data"):
        livro.movimentos(date(2026, 7, 31))


def test_oferta_de_pme_tem_o_ganho_isento_e_fora_do_limite() -> None:
    """IN RFB 1.585/2015 art. 67 I e Lei 13.043/2014 art. 16: ações de PME adquiridas em oferta.
    1.000 a 27,00 (27.000,00) vendidas a 30,00: ganho de 3.000,00 isento, fora do limite de R$ 20
    mil (a venda de 30.000,00 não conta), sem imposto."""
    livro = Livro()
    fim_da_isencao = date(2023, 12, 31)
    livro.evento(date(2019, 6, 17), "AGRO3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "27000",
                 pme_ate=fim_da_isencao)  # fmt: skip
    livro.nota(date(2019, 7, 15), livro.op("V", "AGRO3", 1000, "30"))
    ap = livro.apurar(date(2019, 7, 31))
    assert mes(ap, 2019, 7).isento_pme == D(3000)
    assert categoria(ap, 2019, 7, Categoria.COMUM).imposto == D(0)


def test_bonificacao_noutra_classe_abre_posicao_com_o_custo() -> None:
    """Lei 9.249/1995 art. 10 §1º: o custo é a parcela capitalizada que corresponde ao acionista.
    Quem tem PN recebe 100 ON a 2,00: abre posição de ON com custo de 200,00. Sem a marca, a
    bonificação sem posição continua falhando alto."""
    livro = Livro()
    livro.nota(date(2025, 1, 6), livro.op("C", "DDDD4", 1000, "10"))
    livro.evento(date(2025, 2, 3), "DDDD3", TipoEvento.BONIFICACAO, "C", "100", "200",
                 natureza=NaturezaEvento.OUTRA_CLASSE)  # fmt: skip
    posicoes = {p.ticker: (p.quantidade, p.custo) for p in livro.carteira(date(2025, 2, 28))}
    assert posicoes == {"DDDD4": (D(1000), D(10000)), "DDDD3": (D(100), D(200))}


def test_venda_cobre_so_a_oferta_mais_proxima() -> None:
    """Venda de 100 a 30,00 com duas ofertas na janela, de 150 e de 50: só a primeira é puxada
    para o dia da venda (a 20,00: custo de 3.000,00; baixa de 2.000,00, ganho de 1.000,00 isento).
    A segunda entra na própria data e a posição termina com 100."""
    livro = Livro()
    livro.nota(date(2026, 6, 30), livro.op("V", "CCCC3", 100, "30"))
    livro.evento(date(2026, 7, 1), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "150", "3000")
    livro.evento(date(2026, 7, 3), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "50", "1000")
    ap = livro.apurar(date(2026, 7, 31))
    assert categoria(ap, 2026, 6, Categoria.COMUM).isento == D("1000.00")
    datas = [m.data for m in ap.movimentos if m.origem.startswith("evento:")]
    assert datas == [date(2026, 6, 30), date(2026, 7, 3)]
    [p] = livro.carteira(date(2026, 7, 31))
    assert (p.quantidade, p.custo) == (D(100), D(2000))


def test_compra_no_mesmo_dia_cobre_a_venda_e_nao_puxa_a_oferta() -> None:
    """Compra e venda de 100 no mesmo dia: a posição não fica vendida, a oferta fica na sua data."""
    livro = Livro()
    livro.nota(
        date(2026, 6, 30), livro.op("C", "CCCC3", 100, "25"), livro.op("V", "CCCC3", 100, "26")
    )
    livro.evento(date(2026, 7, 1), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "100", "2000")
    movimentos = livro.movimentos(date(2026, 7, 31))
    assert [m.data for m in movimentos if m.origem.startswith("evento:")] == [date(2026, 7, 1)]


def test_oferta_antecipada_de_um_ano_para_o_outro() -> None:
    """Venda em 30/12/2025 e liquidação em 02/01/2026: o ganho é de dezembro de 2025. 1.000 a 26,00
    (26.000,00), custo de 20.000,00: ganho de 6.000,00, imposto de 900,00."""
    livro = Livro()
    livro.nota(date(2025, 12, 30), livro.op("V", "CCCC3", 1000, "26"))
    livro.evento(date(2026, 1, 2), "CCCC3", TipoEvento.OFERTA_PUBLICA, "C", "1000", "20000")
    dezembro = categoria(livro.apurar(date(2026, 1, 31)), 2025, 12, Categoria.COMUM)
    assert (dezembro.resultado, dezembro.imposto) == (D("6000.00"), D("900.00"))


def test_bonificacao_noutra_classe_com_posicao_vendida_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2025, 1, 6), livro.op("V", "DDDD3", 100, "10"))
    livro.evento(date(2025, 2, 3), "DDDD3", TipoEvento.BONIFICACAO, "C", "50", "100",
                 natureza=NaturezaEvento.OUTRA_CLASSE)  # fmt: skip
    with pytest.raises(MotorError, match="há venda antes desta data"):
        livro.movimentos(date(2025, 2, 28))
