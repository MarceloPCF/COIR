"""Multa e juros de mora da DARF paga em atraso (Lei 9.430/1996 art. 61).

Juros (art. 61 §3º c/c art. 5º §3º): a Selic acumulada de cada mês, da do mês seguinte ao do
vencimento até a do mês anterior ao do pagamento (soma simples, tabela mensal do BCB em
data/selic_mensal.csv), mais 1% no mês do pagamento; pago no mesmo mês do vencimento, sem juros.
Base: o principal (não a multa); centavos truncados (o Sicalc trunca: 12,3457 -> 12,34, conferido
pelo usuário em 17/10/2026). Data não útil vale no próximo dia útil, o que pode mudar o mês.
Selic de mês posterior ao último da tabela: o valor sai sem ela, marcado como estimativa.

- 0,33% por dia de atraso sobre o principal, limitada a 20% (caput e §2º). Os percentuais vêm
  de data/parametros_fiscais.csv, pela vigência na competência (fato gerador, art. 61 caput).
- Contagem (§1º): a lei conta "a partir do primeiro dia subseqüente ao do vencimento"; a
  Receita (Sicalc; página "Como calcular multa de mora") conta do primeiro dia ÚTIL seguinte ao
  vencimento até o pagamento. Parâmetro ``multa_mora_inicio_contagem`` (``rfb`` ou ``lei``);
  decisão do usuário (17/10/2026): ``rfb``, o que a Receita cobra de fato.
- Pagamento em dia não útil vale no próximo dia útil (glossário do Sicalc; decisão do usuário).
- Centavos: a lei é omissa; o Sicalc trunca (``multa_mora_arredondamento = ROUND_DOWN``,
  conferido pelo usuário em 17/10/2026).
- DARF abaixo do mínimo acumulada não tem multa: o vencimento é o do período em que o total
  chega ao mínimo (art. 68 §1º), e é esse vencimento que entra aqui.
"""

from collections.abc import Mapping
from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

from coir.domain.fiscal import ZERO, JurosMora, MultaMora
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


def juros_mora(
    principal: Decimal,
    competencia: date,
    vencimento: date,
    pagamento: date,
    params: ParametrosFiscais,
    cal: Calendario,
    selic: Mapping[tuple[int, int], Decimal],
) -> JurosMora:
    """Juros sobre o `principal` vencido em `vencimento` e pago em `pagamento`. Mês depois do
    último da tabela Selic entra como "sem taxa" (estimativa, avisada); mês faltando no meio da
    tabela ou antes dela falha alto (tabela incompleta)."""
    if principal < 0:
        raise MotorError(f"principal negativo nos juros de mora: {principal}")
    efetivo = proximo_dia_util(pagamento, cal)
    if (efetivo.year, efetivo.month) <= (vencimento.year, vencimento.month):
        return JurosMora(pagamento=efetivo, percentual=ZERO, valor=ZERO)
    ultimo = max(selic, default=(0, 0))
    percentual = params.decimal("juros_mora_mes_pagamento", competencia)
    taxas: list[tuple[str, Decimal]] = []
    faltando: list[str] = []
    ano, mes = vencimento.year, vencimento.month
    while True:
        ano, mes = (ano + 1, 1) if mes == 12 else (ano, mes + 1)
        if (ano, mes) == (efetivo.year, efetivo.month):
            break
        if (ano, mes) in selic:
            taxas.append((f"{mes:02d}/{ano}", selic[(ano, mes)]))
            percentual += selic[(ano, mes)] / 100
        elif (ano, mes) > ultimo:
            faltando.append(f"{mes:02d}/{ano}")
        else:
            raise MotorError(f"Selic de {mes:02d}/{ano} ausente na tabela (data/selic_mensal.csv)")
    modo = params.texto("juros_mora_arredondamento", competencia)
    if modo not in ARREDONDAMENTOS:
        raise MotorError(f"juros_mora_arredondamento inválido: {modo!r} ({list(ARREDONDAMENTOS)})")
    valor = (principal * percentual).quantize(CENTAVO, rounding=ARREDONDAMENTOS[modo])
    return JurosMora(pagamento=efetivo, percentual=percentual, valor=valor, taxas=tuple(taxas),
                     meses_sem_taxa=tuple(faltando))  # fmt: skip
