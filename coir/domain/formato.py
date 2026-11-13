"""Formatação de valores para os textos que o usuário lê (avisos, mensagens)."""

from decimal import Decimal


def reais(valor: Decimal) -> str:
    """Valor em reais no padrão brasileiro: ``Decimal("112626.22")`` vira "R$ 112.626,22"."""
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
