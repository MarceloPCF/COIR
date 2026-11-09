"""Compra a termo em aberto na Carteira: linha própria até a liquidação, fora da posição e do PM
(regra R18 do usuário, F2: "durante o contrato, as ações compradas a termo ficam numa posição
separada, como na B3")."""

from datetime import date

from coir.domain.fiscal import Visao
from coir.engine.movimentos import carteira, termos_em_aberto
from coir.engine.rateio import ratear
from coir.servicos.carteira import posicoes
from tests.fiscal_cases.apoio import D, Livro, parametros


def _livro() -> Livro:
    """1.000 AAAA3 à vista a 25,00 em 02/01/2024; numa nota de 03/01, duas compras a termo de
    AAAA3 (300 a 31,00 e 200 a 30,00, liquidação 04/03) com R$ 40,00 de custos e 100 BBBB3 à vista
    a 10,00. Custos rateados pelo valor: 9.300 + 6.000 + 1.000 = 16.300."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "25"))
    livro.nota(date(2024, 1, 3),
               livro.op("C", "AAAA3", 300, "31", mercado="TERMO", liquidacao=date(2024, 3, 4)),
               livro.op("C", "AAAA3", 200, "30", mercado="TERMO", liquidacao=date(2024, 3, 4)),
               livro.op("C", "BBBB3", 100, "10"), custos="40")  # fmt: skip
    return livro


def test_termo_em_aberto_tem_o_custo_que_entrara_no_pm() -> None:
    livro = _livro()
    termos = termos_em_aberto(livro.notas, date(2024, 2, 5))
    assert [(t.ticker, t.quantidade, t.liquidacao) for t in termos] == [
        ("AAAA3", D(300), date(2024, 3, 4)), ("AAAA3", D(200), date(2024, 3, 4))]  # fmt: skip
    custo = sum((t.custo for t in termos), D(0))
    # mesmo custo que a liquidação põe no PM
    pos = carteira(livro.notas, livro.eventos, parametros(), date(2024, 3, 4))
    aaaa = next(p for p in pos if p.ticker == "AAAA3")
    assert aaaa.custo == D(25000) + custo
    rateio = ratear(D(40), [D(9300), D(6000), D(1000)])
    assert custo == D(15300) + rateio[0] + rateio[1]


def test_termo_some_da_lista_no_dia_da_liquidacao_e_nao_aparece_antes_do_pregao() -> None:
    livro = _livro()
    assert termos_em_aberto(livro.notas, date(2024, 1, 2)) == []
    assert termos_em_aberto(livro.notas, date(2024, 3, 3))
    assert termos_em_aberto(livro.notas, date(2024, 3, 4)) == []


def test_carteira_mostra_o_termo_numa_linha_abaixo_do_ativo_sem_mexer_no_pm() -> None:
    livro = _livro()
    dia = date(2024, 2, 5)
    pos = carteira(livro.notas, livro.eventos, parametros(), dia)
    linhas = posicoes(pos, {1: "123"}, 1, termos_em_aberto(livro.notas, dia), Visao.TITULAR)
    assert [(x.ticker, x.tipo, x.quantidade) for x in linhas] == [
        ("AAAA3", "Ação", D(1000)),
        ("AAAA3", "Termo · liquida em 04/03/2024", D(500)),
        ("BBBB3", "Ação", D(100)),
    ]  # fmt: skip
    vista, termo = linhas[0], linhas[1]
    assert vista.pm == D(25) and vista.liquidacao is None
    assert termo.liquidacao == date(2024, 3, 4) and termo.nota == livro.notas[1].id
    assert termo.custo == sum((t.custo for t in termos_em_aberto(livro.notas, dia)), D(0))


def test_na_visao_por_conta_o_termo_leva_corretora_e_conta() -> None:
    livro = _livro()
    dia = date(2024, 2, 5)
    pos = carteira(livro.notas, livro.eventos, parametros(), dia, Visao.CONTA)
    linhas = posicoes(pos, {1: "123"}, 1, termos_em_aberto(livro.notas, dia), Visao.CONTA)
    termo = next(x for x in linhas if x.liquidacao is not None)
    assert (termo.corretora, termo.conta, termo.numero_conta) == ("XP", 1, "123")
