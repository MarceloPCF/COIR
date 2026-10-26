"""Saldos iniciais de quem começa a usar o COIR sem o histórico de notas (F9, decisões do usuário
de 25/10/2026; parecer do especialista).

Custódia inicial: a posição em 31/12 entra pelo custo de aquisição (IN RFB 1.585/2015 art. 58),
o mesmo do Bens e Direitos, sem resultado e sem efeito no day trade (art. 65 §2º). Prejuízo a
compensar em 31/12, da ficha de renda variável de dezembro: vale em janeiro seguinte, separado em
comum, day trade (arts. 64 e 65) e FII (art. 37 §2º).
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, SaldoInicial, TipoAtivo, TipoEvento
from coir.engine.apuracao import apurar
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS, Livro, categoria

D = Decimal
FIM = date(2024, 12, 31)


def test_custodia_inicial_entra_pelo_custo_e_a_venda_usa_esse_pm() -> None:
    """1.000 ações com custo de 25.000,00 em 31/12/2024 (PM 25,00); venda de 1.000 a 30,00 em
    março/2025 (30.000,00, acima de R$ 20 mil): ganho tributável de 5.000,00."""
    livro = Livro()
    livro.evento(FIM, "PETR4", TipoEvento.CUSTODIA_INICIAL, "C", "1000", "25000")
    livro.nota(date(2025, 3, 10), livro.op("V", "PETR4", 1000, "30"))
    ap = livro.apurar(date(2025, 3, 31))
    assert categoria(ap, 2025, 3, Categoria.COMUM).resultado == D("5000.00")
    assert all(m.resultado == 0 for m in ap.movimentos if m.data == FIM)


def test_prejuizo_inicial_compensa_por_categoria() -> None:
    """Prejuízo em 31/12/2024: comum 3.000,00, day trade 500,00, FII 200,00. Em março/2025 o
    ganho comum de 5.000,00 compensa os 3.000,00 (base 2.000,00, imposto 300,00); o de day trade
    e o de FII seguem a compensar em janeiro (só compensam com a própria categoria)."""
    livro = Livro()
    livro.evento(FIM, "PETR4", TipoEvento.CUSTODIA_INICIAL, "C", "1000", "25000")
    livro.nota(date(2025, 3, 10), livro.op("V", "PETR4", 1000, "30"))
    saldo = SaldoInicial(titular=1, ano=2024, comum=D("3000"), day_trade=D("500"), fii=D("200"))
    ap = apurar(livro.notas, livro.eventos, REFS.parametros, REFS.calendario,
                date(2025, 3, 31), saldos_iniciais=[saldo])  # fmt: skip
    jan = {c: categoria(ap, 2025, 1, c).prejuizo_anterior for c in Categoria}
    assert jan == {Categoria.COMUM: D("-3000"), Categoria.DAY_TRADE: D("-500"),
                   Categoria.FII: D("-200")}  # fmt: skip
    marco = categoria(ap, 2025, 3, Categoria.COMUM)
    assert (marco.base, marco.imposto) == (D("2000.00"), D("300.00"))
    assert categoria(ap, 2025, 3, Categoria.DAY_TRADE).prejuizo_a_compensar == D("-500")


@pytest.mark.parametrize(
    ("cv", "tipo_ativo"), [("V", TipoAtivo.ACAO), ("C", TipoAtivo.OPCAO), ("C", TipoAtivo.FUTURO)]
)
def test_custodia_inicial_so_de_posicao_comprada_a_vista(cv: str, tipo_ativo: TipoAtivo) -> None:
    """Decisão do usuário: opções, futuros e posição vendida pedem as notas desde a abertura."""
    livro = Livro()
    livro.evento(FIM, "XPTO3", TipoEvento.CUSTODIA_INICIAL, cv, "100", "1000",
                 tipo_ativo=tipo_ativo)  # fmt: skip
    livro.nota(date(2025, 1, 10), livro.op("C", "PETR4", 100, "30"))
    with pytest.raises(MotorError, match="custódia inicial"):
        livro.apurar(date(2025, 1, 31))
