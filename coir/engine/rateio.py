"""Rateio dos custos da nota pelo valor das operações (IN RFB 1.585/2015 art. 58 §2º: custos de
corretagem e taxas integram o custo de aquisição e reduzem o valor de alienação).

Rateio em centavos pelo maior resto: as parcelas somam exatamente o total da nota.
"""

from collections.abc import Sequence
from decimal import ROUND_DOWN, Decimal

from coir.domain.fiscal import ZERO
from coir.erros import MotorError

CENTAVO = Decimal("0.01")


def ratear(total: Decimal, pesos: Sequence[Decimal]) -> list[Decimal]:
    """Divide `total` (em reais, 2 casas) na proporção dos pesos (valores absolutos)."""
    if total < 0 or total != total.quantize(CENTAVO):
        raise MotorError(f"custo da nota inválido para rateio: {total}")
    absolutos = [abs(p) for p in pesos]
    soma = sum(absolutos, ZERO)
    if not absolutos:
        if total != 0:
            raise MotorError(f"custo de {total} numa nota sem operações para ratear")
        return []
    if soma == 0:
        absolutos = [Decimal(1)] * len(absolutos)  # sem valor (ex.: só AJUPOS): partes iguais
        soma = Decimal(len(absolutos))
    exatas = [total * p / soma for p in absolutos]
    parcelas = [e.quantize(CENTAVO, rounding=ROUND_DOWN) for e in exatas]
    sobra = int((total - sum(parcelas, ZERO)) / CENTAVO)
    # centavos que faltam vão para os maiores restos (empate: a operação de menor índice)
    ordem = sorted(range(len(exatas)), key=lambda i: (-(exatas[i] - parcelas[i]), i))
    for i in ordem[:sobra]:
        parcelas[i] += CENTAVO
    return parcelas
