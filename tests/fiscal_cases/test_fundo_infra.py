"""FI-Infra (Lei 12.431/2011 art. 3º) e FIP-IE (Lei 11.478/2007 art. 2º §1º I), tipo FUNDO_INFRA
pela lista de fundos da B3: o ganho de pessoa física tem alíquota zero, fica fora da apuração
mensal, do limite de R$ 20 mil e do DARF, e o prejuízo não compensa outros ganhos (IN RFB
1.585/2015 art. 64 só admite as perdas dos arts. 27, 58 e 60 a 62). Parecer do especialista de
09/11/2026, seguido pelo usuário."""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo
from tests.fiscal_cases.apoio import D, Livro, categoria

INFRA = TipoAtivo.FUNDO_INFRA


def test_ganho_de_fi_infra_fica_fora_da_apuracao_e_do_limite() -> None:
    """Fevereiro: ações 19.000 (isentas) + JURO11 11.000; sem o fundo na soma, as ações seguem
    isentas, e o ganho de 1.000 do fundo vai para a lista à parte, sem imposto."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "PETR4", 1000, "10"),
               livro.op("C", "JURO11", 100, "100", tipo=INFRA))  # fmt: skip
    livro.nota(date(2025, 2, 10), livro.op("V", "PETR4", 1000, "19"),
               livro.op("V", "JURO11", 100, "110", tipo=INFRA))  # fmt: skip
    ap = livro.apurar(date(2025, 2, 28))
    fevereiro = categoria(ap, 2025, 2, Categoria.COMUM)
    assert (fevereiro.vendas_acoes, fevereiro.isento, fevereiro.imposto) == (D(19000), D(9000), 0)
    [fundo] = ap.exclusiva
    assert (fundo.ticker, fundo.resultado, fundo.tipo) == ("JURO11", D(1000), INFRA)


def test_prejuizo_de_fip_ie_nao_compensa_ganho_de_acoes() -> None:
    """Prejuízo de 5.000 no FIP-IE em março não reduz o ganho tributável de 10.000 em abril."""
    livro = Livro()
    livro.nota(date(2025, 3, 3), livro.op("C", "XPIE11", 100, "100", tipo=INFRA),
               livro.op("C", "PETR4", 1000, "20"))  # fmt: skip
    livro.nota(date(2025, 3, 20), livro.op("V", "XPIE11", 100, "50", tipo=INFRA))
    livro.nota(date(2025, 4, 10), livro.op("V", "PETR4", 1000, "30"))
    ap = livro.apurar(date(2025, 4, 30))
    assert categoria(ap, 2025, 3, Categoria.COMUM).prejuizo_a_compensar == 0
    assert categoria(ap, 2025, 4, Categoria.COMUM).imposto == D("1500.00")
    assert [e.resultado for e in ap.exclusiva] == [D(-5000)]


def test_day_trade_de_fi_infra_tambem_fica_fora() -> None:
    livro = Livro()
    livro.nota(date(2025, 5, 5), livro.op("C", "JURO11", 100, "100", tipo=INFRA),
               livro.op("V", "JURO11", 100, "101", tipo=INFRA))  # fmt: skip
    ap = livro.apurar(date(2025, 5, 31))
    assert categoria(ap, 2025, 5, Categoria.DAY_TRADE).imposto == 0
    assert [e.resultado for e in ap.exclusiva] == [D(100)]
