"""Leitura de valores digitados na tela (formato brasileiro, tolerante ao ponto decimal)."""

from decimal import Decimal, InvalidOperation


def valor_br(texto: str) -> Decimal | None:
    """ "1.234,56", "1234,56", "1234.56" e "523.47" -> Decimal; "1.234" (milhar) -> 1234.
    Com vírgula, o ponto é separador de milhar; sem vírgula, um único ponto seguido de até duas
    casas é o decimal. Texto que não é número -> None."""
    t = texto.replace("R$", "").replace(" ", "").strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif not (t.count(".") == 1 and len(t.split(".")[1]) <= 2):
        t = t.replace(".", "")
    try:
        valor = Decimal(t)
    except InvalidOperation:
        return None
    return valor if valor.is_finite() else None
