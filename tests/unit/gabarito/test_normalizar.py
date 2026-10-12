from datetime import date
from decimal import Decimal

import pytest

from coir.erros import GabaritoError
from coir.gabarito import normalizar as nz
from tests.unit.gabarito.fabrica import CPF_A, CPF_A_ERRADO, com_dv, mascara


def test_para_decimal_remove_ruido_do_float() -> None:
    assert nz.para_decimal(26.919999999999998) == Decimal("26.92000000")
    assert nz.para_decimal(62.28973333333333) == Decimal("62.28973333")
    assert nz.para_decimal(None) is None
    assert nz.para_decimal("  ") is None
    assert nz.para_decimal("1,5") == Decimal("1.50000000")


@pytest.mark.parametrize("valor", [True, "abc", object()])
def test_para_decimal_rejeita(valor: object) -> None:
    with pytest.raises(GabaritoError):
        nz.para_decimal(valor)


def test_data_excel() -> None:
    # 39556 = 18/04/2008 (nota Ágora com liquidação em 24/04, após o feriado de 21/04)
    assert nz.data_excel(39556.0) == date(2008, 4, 18)
    assert nz.data_excel(None) is None
    with pytest.raises(GabaritoError):
        nz.data_excel("18/04/2008")


@pytest.mark.parametrize(
    ("bruto", "esperado"), [("XP ", "XP"), ("Agora", "AGORA"), ("Ágora", "AGORA"), ("btg", "BTG")]
)
def test_corretora(bruto: str, esperado: str) -> None:
    assert nz.corretora(bruto) == esperado


def test_texto_inteiro() -> None:
    assert nz.texto_inteiro(2588033.0) == "2588033"
    assert nz.texto_inteiro(None) is None
    with pytest.raises(GabaritoError):
        nz.texto_inteiro(1.5)


def test_corrigir_cpf_valido_passa_direto() -> None:
    assert nz.corrigir_cpf(mascara(CPF_A), {CPF_A}) == (CPF_A, False)


def test_corrigir_cpf_digitado_errado() -> None:
    assert nz.corrigir_cpf(mascara(CPF_A_ERRADO), {CPF_A}) == (CPF_A, True)


def test_corrigir_cpf_ambiguo_ou_sem_candidato_falha_sem_expor_cpf() -> None:
    # mesmos 9 primeiros dígitos do CPF errado + DV correto: válido e a ≤ 2 dígitos dele
    outro = com_dv(CPF_A_ERRADO[:9])
    with pytest.raises(GabaritoError) as erro:
        nz.corrigir_cpf(CPF_A_ERRADO, {CPF_A, outro})
    assert CPF_A_ERRADO not in str(erro.value)
    with pytest.raises(GabaritoError):
        nz.corrigir_cpf(CPF_A_ERRADO, set())
