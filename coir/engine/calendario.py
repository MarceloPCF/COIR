"""Dias úteis, vencimento de DARF e de opções, liquidação de termo (datas sempre por parâmetro).

- DARF: último dia útil (bancário) do mês seguinte ao da apuração (RIR/2018 art. 915;
  Lei 9.430/1996 art. 68 §1º para o acúmulo abaixo do mínimo).
- Opções sobre ações: regra de data/vencimentos_opcoes.csv com vigência (hoje: 3ª sexta-feira do
  mês; sem pregão, o pregão anterior), com exceções remarcadas pela B3.
- Termo: liquidação na data do pregão + prazo do contrato (dias corridos), no pregão seguinte se
  cair em dia sem pregão.
"""

from datetime import date, timedelta

from coir.domain.parametros import Calendario
from coir.erros import ReferenciaError

UM_DIA = timedelta(days=1)


def dia_util(dia: date, cal: Calendario) -> bool:
    cal.cobre(dia)
    return dia.weekday() < 5 and dia not in cal.nao_uteis


def dia_de_pregao(dia: date, cal: Calendario) -> bool:
    return dia_util(dia, cal) and dia not in cal.sem_pregao


def fim_do_mes(ano: int, mes: int) -> date:
    return (date(ano + mes // 12, mes % 12 + 1, 1)) - UM_DIA


def ultimo_dia_util(ano: int, mes: int, cal: Calendario) -> date:
    dia = fim_do_mes(ano, mes)
    while not dia_util(dia, cal):
        dia -= UM_DIA
    return dia


def vencimento_darf(ano: int, mes: int, cal: Calendario) -> date:
    """Mês de apuração (ano, mes) -> último dia útil do mês seguinte."""
    return ultimo_dia_util(ano + mes // 12, mes % 12 + 1, cal)


def _pregao(dia: date, passo: timedelta, cal: Calendario) -> date:
    while not dia_de_pregao(dia, cal):
        dia += passo
    return dia


def vencimento_opcao(ano: int, mes: int, cal: Calendario, semana: int | None = None) -> date:
    """Data de vencimento da série do mês (semanal: `semana` = n-ésima ocorrência do dia)."""
    chave = f"{ano:04d}-{mes:02d}" if semana is None else f"{ano:04d}-{mes:02d}-W{semana}"
    if chave in cal.excecoes_vencimento:
        return cal.excecoes_vencimento[chave]
    referencia = date(ano, mes, 15)
    regras = [r for r in cal.vencimentos if r.vigente(referencia)]
    if len(regras) != 1:
        raise ReferenciaError(
            f"regra de vencimento de opções {'ausente' if not regras else 'repetida'} para "
            f"{mes:02d}/{ano} (data/vencimentos_opcoes.csv)"
        )
    regra = regras[0]
    ordem = regra.ordem if semana is None else semana
    primeiro = date(ano, mes, 1)
    dia = primeiro + timedelta(days=(regra.dia_semana - primeiro.weekday()) % 7)
    dia += timedelta(weeks=ordem - 1)
    if dia.month != mes:
        raise ReferenciaError(f"vencimento de opções: sem {ordem}ª ocorrência em {mes:02d}/{ano}")
    if regra.ajuste not in ("anterior", "seguinte"):
        raise ReferenciaError(f"ajuste de vencimento inválido: {regra.ajuste!r}")
    return _pregao(dia, -UM_DIA if regra.ajuste == "anterior" else UM_DIA, cal)


def liquidacao_termo(pregao: date, prazo_dias: int, cal: Calendario) -> date:
    return _pregao(pregao + timedelta(days=prazo_dias), UM_DIA, cal)
