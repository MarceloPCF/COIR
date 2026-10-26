"""Day trade pela marcação D da corretora (F9, decisão do usuário de 25/10/2026).

IN RFB 1.585/2015 art. 65 §3º: o day trade casa "pela ordem" a 1ª compra com a 1ª venda, isto
é, na ordem cronológica dos negócios. A nota não traz o horário; a corretora, que o tem e retém o
IRRF de 1% (art. 65 §5º), marca os negócios de day trade com "D". O COIR usa a marcação quando ela
é consistente (quantidade D de compra = de venda = a menor das duas no dia) e, sem marcação ou
com marcação inconsistente, a ordem impressa (parâmetro day_trade_criterio).
"""

from datetime import date
from decimal import Decimal

from coir.domain.fiscal import Categoria, OperacaoMotor
from tests.fiscal_cases.apoio import Livro, categoria, parametros

D = Decimal
DIA = date(2019, 7, 24)


def _marcada(o: OperacaoMotor) -> OperacaoMotor:
    return o.model_copy(update={"marcada_dt": True})


def _livro(marcacao_consistente: bool = True) -> Livro:
    """Caso da AMER3 de 24/07/2019: compra 200 a 36,19 (sem D), compra 300 a 36,09 (D) e venda
    300 a 35,63 (D), sem posição antes. Em agosto, venda das 200 que sobraram a 37,00."""
    livro = Livro()
    compra_d = livro.op("C", "AMER3", 300, "36.09")
    livro.nota(
        DIA,
        livro.op("C", "AMER3", 200, "36.19"),
        _marcada(compra_d) if marcacao_consistente else compra_d,
        _marcada(livro.op("V", "AMER3", 300, "35.63")),
    )
    livro.nota(date(2019, 8, 15), livro.op("V", "AMER3", 200, "37"))
    return livro


def test_marcacao_d_consistente_forma_o_day_trade() -> None:
    """Pela marcação: day trade = 300 x 35,63 - 300 x 36,09 = -138,00; sobram 200 a 36,19, e a
    venda de agosto dá 200 x (37,00 - 36,19) = +162,00."""
    ap = _livro().apurar(date(2019, 8, 31))
    assert categoria(ap, 2019, 7, Categoria.DAY_TRADE).resultado == D("-138.00")
    assert categoria(ap, 2019, 8, Categoria.COMUM).isento == D("162.00")  # vendas < R$ 20 mil
    assert not ap.avisos


def test_parametro_ordem_impressa_mantem_o_criterio_anterior() -> None:
    """Pela ordem impressa: day trade = 10.689,00 - (200 x 36,19 + 100 x 36,09) = -158,00; sobram
    200 a 36,09, e agosto dá +182,00. O total dos dois meses é igual nos dois critérios (+24,00)."""
    ap = _livro().apurar(date(2019, 8, 31), parametros(day_trade_criterio="ordem_impressa"))
    assert categoria(ap, 2019, 7, Categoria.DAY_TRADE).resultado == D("-158.00")
    assert categoria(ap, 2019, 8, Categoria.COMUM).isento == D("182.00")


def test_marcacao_inconsistente_volta_a_ordem_com_aviso() -> None:
    """Só a venda marcada D (300) e nenhuma compra: a marcação não fecha com o dia; vale a ordem
    impressa, com aviso."""
    ap = _livro(marcacao_consistente=False).apurar(date(2019, 8, 31))
    assert categoria(ap, 2019, 7, Categoria.DAY_TRADE).resultado == D("-158.00")
    assert any("marcação D da nota não fecha" in a for a in ap.avisos)
