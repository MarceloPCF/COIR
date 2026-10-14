"""Validação de CPF (só dígitos verificadores; nenhum CPF real fica no código)."""

import re


def so_digitos(texto: str) -> str:
    """Remove máscara (pontos, traço, espaços) de um CPF."""
    return re.sub(r"\D", "", texto)


def cpf_valido(cpf: str) -> bool:
    """Confere os dígitos verificadores de um CPF (com ou sem máscara)."""
    digitos = so_digitos(cpf)
    if len(digitos) != 11 or len(set(digitos)) == 1:
        return False
    nums = [int(c) for c in digitos]
    for pos in (9, 10):
        soma = sum(n * peso for n, peso in zip(nums[:pos], range(pos + 1, 1, -1), strict=True))
        if (soma * 10) % 11 % 10 != nums[pos]:
            return False
    return True
