"""Outros custos dedutíveis lançados pelo usuário (ex.: corretagem e tarifas do BTC): reduzem o
resultado da categoria no mês, como a aba BTC_Outros da planilha v2.02 (regra do projeto:
docs/regras-fiscais.md, "Outros custos dedutíveis"). A remuneração paga ao doador não é dedutível
por pessoa física (SC Cosit 42/2020); corretagem e emolumentos do tomador: Lei 13.043/2014 art. 9º
par. único, IN RFB 1.585/2015 art. 76 par. único e, desde 2026, art. 56 §3º."""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, OutroCusto
from coir.engine.apuracao import apurar
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS, Livro, categoria, mes

D = Decimal


def _apurar(livro: Livro, *custos: OutroCusto):  # type: ignore[no-untyped-def]
    return apurar(livro.notas, livro.eventos, REFS.parametros, REFS.calendario,
                  date(2024, 3, 31), custos)  # fmt: skip


def test_outro_custo_reduz_o_resultado_comum_do_mes() -> None:
    """Units TAEE11: compra 100 x 40 e venda 100 x 50 em fev/2024 -> ganho 1.000 (tributável:
    units fora da isenção). Custo de BTC de 120,00 em fev: resultado 880,00, imposto 15% = 132,00.
    O custo aparece como movimento (origem custo:1) na memória do mês."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "TAEE11", 100, "40"))
    livro.nota(date(2024, 2, 1), livro.op("V", "TAEE11", 100, "50"))
    custo = OutroCusto(id=1, titular=1, ano=2024, mes=2, categoria=Categoria.COMUM,
                       valor=D("120.00"), descricao="corretagem do BTC")  # fmt: skip
    ap = _apurar(livro, custo)
    c = categoria(ap, 2024, 2, Categoria.COMUM)
    assert (c.resultado, c.imposto) == (D("880.00"), D("132.00"))
    [m] = [m for m in ap.movimentos if m.origem == "custo:1"]
    assert m.resultado == D("-120.00") and m.data == date(2024, 2, 29) and "corretagem" in m.regra


def test_outro_custo_de_day_trade_e_mes_sem_operacao_vira_prejuizo() -> None:
    """Custo de day trade de 50,00 em mar/2024 sem operação no mês: prejuízo de day trade a
    compensar de -50,00."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "TAEE11", 100, "40"))
    custo = OutroCusto(id=2, titular=1, ano=2024, mes=3, categoria=Categoria.DAY_TRADE,
                       valor=D("50.00"))  # fmt: skip
    c = categoria(_apurar(livro, custo), 2024, 3, Categoria.DAY_TRADE)
    assert (c.resultado, c.prejuizo_a_compensar) == (D("-50.00"), D("-50.00"))


def test_outro_custo_invalido_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "TAEE11", 100, "40"))
    with pytest.raises(MotorError, match="outro custo"):
        _apurar(livro, OutroCusto(id=3, titular=1, ano=2024, mes=2, categoria=Categoria.FII,
                                  valor=D(1)))  # fmt: skip


def test_btc_entra_a_corretagem_e_a_tarifa_nao_a_remuneracao() -> None:
    """Venda a descoberto em 10/02/2026: 1.000 x 30,00, custos 10,00, líquido 29.990,00, IRRF 1,50.
    Recompra em 10/03/2026: 1.000 x 25,00, custos 8,00, total 25.008,00. Custos do BTC: remuneração
    150,00 (fora), tarifa 12,00 e corretagem 20,00, lançados em março: 32,00. Resultado 4.982,00 -
    32,00 = 4.950,00, acima do limite (venda de 30.000,00); imposto 742,50; DARF 741,00, vencimento
    30/04/2026. Com os 182,00 o imposto seria 720,00: 22,50 a menos."""
    livro = Livro()
    livro.nota(date(2026, 2, 10), livro.op("V", "PETR4", 1000, "30"), custos="10", irrf="1.50")
    livro.nota(date(2026, 3, 10), livro.op("C", "PETR4", 1000, "25"), custos="8")
    custo = OutroCusto(id=1, titular=1, ano=2026, mes=3, categoria=Categoria.COMUM,
                       valor=D("32.00"), descricao="corretagem e tarifa do BTC")  # fmt: skip
    ap = apurar(livro.notas, livro.eventos, REFS.parametros, REFS.calendario, date(2026, 3, 31),
                (custo,))  # fmt: skip
    marco = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (marco.resultado, marco.imposto) == (D("4950.00"), D("742.50"))
    apurado = mes(ap, 2026, 3)
    assert (apurado.darf, apurado.vencimento) == (D("741.00"), date(2026, 4, 30))


def _a_descoberto_isento(valor_custo: str) -> object:
    livro = Livro()
    livro.nota(date(2026, 2, 10), livro.op("V", "PETR4", 500, "30"))
    livro.nota(date(2026, 3, 10), livro.op("C", "PETR4", 500, "20"))
    custo = OutroCusto(id=1, titular=1, ano=2026, mes=3, categoria=Categoria.COMUM,
                       valor=D(valor_custo))  # fmt: skip
    ap = apurar(livro.notas, livro.eventos, REFS.parametros, REFS.calendario, date(2026, 3, 31),
                (custo,))  # fmt: skip
    return categoria(ap, 2026, 3, Categoria.COMUM)


def test_custo_em_mes_de_ganho_isento_reduz_o_isento() -> None:
    """Venda a descoberto de 500 x 30,00 (15.000,00, dentro do limite) em fevereiro; recompra de 500
    x 20,00 em março: ganho isento de 5.000,00. Custo de 30,00: isento de 4.970,00, sem prejuízo a
    compensar (decisão do usuário de 08/11/2026)."""
    c = _a_descoberto_isento("30")
    assert (c.isento, c.resultado, c.prejuizo_a_compensar) == (D("4970.00"), D(0), D(0))  # type: ignore[attr-defined]


def test_custo_maior_que_o_ganho_isento_vira_prejuizo_so_no_que_sobra() -> None:
    """Custo de 6.000,00 sobre o isento de 5.000,00: isento zero e prejuízo de 1.000,00."""
    c = _a_descoberto_isento("6000")
    assert (c.isento, c.resultado, c.prejuizo_a_compensar) == (D(0), D("-1000.00"), D("-1000.00"))  # type: ignore[attr-defined]
