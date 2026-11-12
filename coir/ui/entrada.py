"""Leitura de valores digitados na tela (formato brasileiro, tolerante ao ponto decimal)."""

import re
from decimal import Decimal, InvalidOperation

NUMERO = re.compile(r"[+-]?(\d+(\.\d*)?|\.\d+)")  # "5", "5.", ".5", "+5", "-5"
LIMITE = Decimal(10) ** 15  # acima disso é engano de digitação (e estoura o arredondamento)


def valor_br(texto: str) -> Decimal | None:
    """ "1.234,56", "1234,56", "1234.56" e "523.47" -> Decimal; "1.234" (milhar) -> 1234.
    Com vírgula, o ponto é separador de milhar; sem vírgula, um único ponto seguido de até duas
    casas é o decimal. Texto que não é número -> None, inclusive notação científica ("1e400"),
    "NaN", "Infinity" e "1_000", que o Decimal aceitaria, e valor de 10^15 para cima."""
    t = texto.replace("R$", "").replace(" ", "").strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    elif not (t.count(".") == 1 and len(t.split(".")[1]) <= 2):
        t = t.replace(".", "")
    if not NUMERO.fullmatch(t):
        return None
    try:
        valor = Decimal(t)
    except InvalidOperation:
        return None
    return valor if abs(valor) < LIMITE else None
