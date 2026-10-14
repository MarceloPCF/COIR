from coir.domain.cpf import cpf_valido, so_digitos
from tests.unit.gabarito.fabrica import CPF_A, CPF_A_ERRADO, mascara


def test_cpf_valido_com_e_sem_mascara() -> None:
    assert cpf_valido(CPF_A)
    assert cpf_valido(mascara(CPF_A))


def test_cpf_invalido() -> None:
    assert not cpf_valido(CPF_A_ERRADO)
    assert not cpf_valido("0" * 11)
    assert not cpf_valido("123")


def test_so_digitos() -> None:
    assert so_digitos(mascara(CPF_A)) == CPF_A
