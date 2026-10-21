"""Validação de CPF e CNPJ (só dígitos verificadores; nenhum CPF real fica no código)."""

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


def cnpj_valido(cnpj: str) -> bool:
    """Confere os dígitos verificadores de um CNPJ (com ou sem máscara; módulo 11)."""
    digitos = so_digitos(cnpj)
    if len(digitos) != 14 or len(set(digitos)) == 1:
        return False
    nums = [int(c) for c in digitos]
    for pos in (12, 13):
        pesos = [*range(pos - 7, 1, -1), *range(9, 1, -1)]
        resto = sum(n * peso for n, peso in zip(nums[:pos], pesos, strict=True)) % 11
        if nums[pos] != (0 if resto < 2 else 11 - resto):
            return False
    return True


def formatar_cnpj(cnpj: str) -> str:
    """ "11728688000147" -> "11.728.688/0001-47" (como o programa IRPF mostra)."""
    d = so_digitos(cnpj)
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
