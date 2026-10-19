"""Outros custos dedutíveis lançados pelo usuário (ex.: taxas de BTC): reduzem o resultado da
categoria no mês, como a aba BTC_Outros da planilha v2.02 (regra do projeto: docs/regras-fiscais.md,
"Outros custos dedutíveis")."""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, OutroCusto
from coir.engine.apuracao import apurar
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS, Livro, categoria

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
                       valor=D("120.00"), descricao="aluguel de ações")  # fmt: skip
    ap = _apurar(livro, custo)
    c = categoria(ap, 2024, 2, Categoria.COMUM)
    assert (c.resultado, c.imposto) == (D("880.00"), D("132.00"))
    [m] = [m for m in ap.movimentos if m.origem == "custo:1"]
    assert m.resultado == D("-120.00") and m.data == date(2024, 2, 29) and "aluguel" in m.regra


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
