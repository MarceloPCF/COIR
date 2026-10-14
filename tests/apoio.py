"""Utilidades compartilhadas pelos testes."""


def cpf_de_teste(base: str = "529982247") -> str:
    """CPF sintético com dígitos verificadores calculados (nenhum CPF real no código)."""
    nums = [int(c) for c in base]
    for pos in (9, 10):
        soma = sum(n * p for n, p in zip(nums, range(pos + 1, 1, -1), strict=False))
        nums.append(soma * 10 % 11 % 10)
    return "".join(map(str, nums))
