"""Rateio dos custos e posição pelo custo médio (unit + propriedades)."""

from decimal import Decimal

import pytest
from hypothesis import given
from hypothesis import strategies as st

from coir.engine.posicao import Posicao
from coir.engine.rateio import ratear
from coir.erros import MotorError

D = Decimal
valores = st.decimals(min_value=D("0.01"), max_value=D("1e7"), places=2)


def test_rateio_proporcional_e_soma_exata() -> None:
    assert ratear(D("10.00"), [D(1), D(1), D(1)]) == [D("3.34"), D("3.33"), D("3.33")]
    assert ratear(D("7.50"), [D(300), D(100)]) == [D("5.63"), D("1.87")]


def test_rateio_sem_custo_ou_sem_operacao() -> None:
    assert ratear(D(0), [D(10), D(20)]) == [D(0), D(0)]
    assert ratear(D(0), []) == []
    with pytest.raises(MotorError):
        ratear(D("1.00"), [])
    with pytest.raises(MotorError):
        ratear(D("-1.00"), [D(1)])


def test_rateio_de_pesos_zerados_divide_igual() -> None:
    assert ratear(D("1.00"), [D(0), D(0), D(0)]) == [D("0.34"), D("0.33"), D("0.33")]


@given(valores, st.lists(valores, min_size=1, max_size=30))
def test_parcelas_do_rateio_somam_o_total(total: Decimal, pesos: list[Decimal]) -> None:
    parcelas = ratear(total, pesos)
    assert sum(parcelas) == total
    assert all(p >= 0 and p == p.quantize(D("0.01")) for p in parcelas)


def test_posicao_comprada_vendida_zerar_e_inverter() -> None:
    p = Posicao()
    p.comprar(D(100), D(1000))
    p.comprar(D(100), D(1200))
    assert p.pm == D(11)
    b = p.vender(D(50), D(600))
    assert (b.resultado, b.custo_baixa, p.pm) == (D(50), D(550), D(11))
    b = p.vender(D(250), D(3000))  # zera 150 e vende 100 a descoberto
    assert (b.resultado, p.quantidade, p.custo) == (D(150), D(-100), D(-1200))
    b = p.comprar(D(150), D(1500))  # cobre 100 (receita 1.200, custo 1.000) e abre 50
    assert (b.resultado, p.quantidade, p.custo) == (D(200), D(50), D(500))
    assert p.virar_po().resultado == D(-500) and p.quantidade == 0 and p.custo == 0


@given(st.lists(st.tuples(st.integers(1, 1000), valores), min_size=1, max_size=20), valores)
def test_comprar_e_vender_tudo_zera_e_resultado_fecha(
    compras: list[tuple[int, Decimal]], venda: Decimal
) -> None:
    """Comprar e vender tudo: quantidade 0, custo 0 e resultado = vendas - compras."""
    p = Posicao()
    for q, v in compras:
        p.comprar(D(q), v)
    total_q = sum((D(q) for q, _ in compras), D(0))
    resultado = p.vender(total_q, venda).resultado
    assert (p.quantidade, p.custo) == (0, 0)
    assert resultado == venda - sum(v for _, v in compras)
