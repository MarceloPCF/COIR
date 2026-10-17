"""Multa de mora da DARF paga em atraso (Lei 9.430/1996 art. 61). Juros Selic: F7.

- 0,33% por dia de atraso sobre o principal, limitada a 20% (caput e §2º). Os percentuais vêm
  de data/parametros_fiscais.csv, pela vigência na competência (fato gerador, art. 61 caput).
- Contagem (§1º): a lei conta "a partir do primeiro dia subseqüente ao do vencimento"; a
  Receita (Sicalc; página "Como calcular multa de mora") conta do primeiro dia ÚTIL seguinte ao
  vencimento até o pagamento. Parâmetro ``multa_mora_inicio_contagem`` (``rfb`` ou ``lei``);
  decisão do usuário (17/10/2026): ``rfb``, o que a Receita cobra de fato.
- Pagamento em dia não útil vale no próximo dia útil (glossário do Sicalc; decisão do usuário).
- Centavos: a lei é omissa; ``multa_mora_arredondamento`` (ROUND_DOWN, decisão do usuário até
  conferir no Sicalc).
- DARF abaixo do mínimo acumulada não tem multa: o vencimento é o do período em que o total
  chega ao mínimo (art. 68 §1º), e é esse vencimento que entra aqui.
"""

from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

from coir.domain.fiscal import ZERO, MultaMora
from coir.domain.parametros import Calendario, ParametrosFiscais
from coir.engine.calendario import UM_DIA, dia_util
from coir.erros import MotorError

CENTAVO = Decimal("0.01")
ARREDONDAMENTOS = {"ROUND_DOWN": ROUND_DOWN, "ROUND_HALF_UP": ROUND_HALF_UP}


def proximo_dia_util(dia: date, cal: Calendario) -> date:
    """O próprio dia, se útil; senão o próximo dia útil (bancário)."""
    while not dia_util(dia, cal):
        dia += UM_DIA
    return dia


def multa_mora(
    principal: Decimal,
    competencia: date,
    vencimento: date,
    pagamento: date,
    params: ParametrosFiscais,
    cal: Calendario,
) -> MultaMora:
    """Multa sobre o `principal` vencido em `vencimento` e pago em `pagamento`."""
    if principal < 0:
        raise MotorError(f"principal negativo na multa de mora: {principal}")
    efetivo = proximo_dia_util(pagamento, cal)
    if efetivo <= vencimento:
        dias = 0
    else:
        inicio = params.texto("multa_mora_inicio_contagem", competencia)
        if inicio == "lei":
            dias = (efetivo - vencimento).days
        elif inicio == "rfb":
            dias = (efetivo - proximo_dia_util(vencimento + UM_DIA, cal)).days + 1
        else:
            raise MotorError(f"multa_mora_inicio_contagem inválido: {inicio!r} (lei ou rfb)")
    modo = params.texto("multa_mora_arredondamento", competencia)
    if modo not in ARREDONDAMENTOS:
        raise MotorError(f"multa_mora_arredondamento inválido: {modo!r} ({list(ARREDONDAMENTOS)})")
    percentual = min(dias * params.decimal("multa_mora_dia", competencia),
                     params.decimal("multa_mora_teto", competencia))  # fmt: skip
    valor = (principal * percentual).quantize(CENTAVO, rounding=ARREDONDAMENTOS[modo])
    return MultaMora(pagamento=efetivo, ajustado=efetivo != pagamento, dias=dias,
                     percentual=percentual if dias else ZERO, valor=valor)  # fmt: skip
