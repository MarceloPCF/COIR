from datetime import date
from decimal import Decimal

import pytest
from hypothesis import given
from hypothesis import strategies as st

from coir.domain.nota import Montante
from coir.erros import ExtracaoError
from coir.parsing.numeros import (
    compactar,
    data_br,
    decimal_br,
    montante,
    montantes,
    valor_apos_rotulo,
)


@pytest.mark.parametrize(
    ("texto", "esperado"),
    [
        ("1.234,56", "1234.56"),
        ("1234,56", "1234.56"),
        ("1800", "1800"),
        ("1.400", "1400"),
        ("97.810,0000", "97810.0000"),
        ("-2", "-2"),
    ],
)
def test_decimal_br(texto: str, esperado: str) -> None:
    assert decimal_br(texto) == Decimal(esperado)


def test_decimal_br_invalido() -> None:
    with pytest.raises(ExtracaoError):
        decimal_br("12a,00")


def test_montante_com_barra_e_dc_variaveis() -> None:
    assert montantes("0,00| 0,41 1,60 0,56 0,48 | D") == [
        Montante(valor=Decimal("0.00")),
        Montante(valor=Decimal("0.41")),
        Montante(valor=Decimal("1.60")),
        Montante(valor=Decimal("0.56")),
        Montante(valor=Decimal("0.48"), dc="D"),
    ]
    assert montantes("0,00 0,08 0,00 | 44,00 | C 2,64 | D")[3] == Montante(
        valor=Decimal("44.00"), dc="C"
    )
    assert montante("0,00|D") == Montante(valor=Decimal("0.00"), dc="D")
    assert montante("1,49 D").assinado == Decimal("-1.49")


def test_rotulos_colados_do_layout_antigo() -> None:
    assert compactar("Taxa deliquidação") == compactar("Taxa de liquidação")
    assert valor_apos_rotulo("Taxa deliquidação 1,49 D", "Taxa de liquidação") == Montante(
        valor=Decimal("1.49"), dc="D"
    )
    # rótulo sem valor na linha (cabeçalho de seção) é ignorado; vale a próxima ocorrência
    texto = "Corretagem / Despesas\nTotal corretagem / Despesas 20,40\nCorretagem 20,00 D"
    assert valor_apos_rotulo(texto, "Corretagem") == Montante(valor=Decimal("20.00"), dc="D")


def test_data_br() -> None:
    assert data_br("18/04/2008") == date(2008, 4, 18)
    with pytest.raises(ExtracaoError):
        data_br("31/02/2008")


@given(st.decimals(min_value=0, max_value=10**9, places=2, allow_nan=False))
def test_propriedade_formato_brasileiro_ida_e_volta(valor: Decimal) -> None:
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    assert decimal_br(texto) == valor
