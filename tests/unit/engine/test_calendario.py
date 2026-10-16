"""Calendário: dias úteis, DARF, vencimento de opções e liquidação de termo (data/)."""

from datetime import date

import pytest

from coir import referencias
from coir.engine.calendario import (
    dia_de_pregao,
    dia_util,
    liquidacao_termo,
    ultimo_dia_util,
    vencimento_darf,
    vencimento_opcao,
)
from coir.erros import ReferenciaError
from tools.gerar_feriados import pascoa

CAL = referencias.carregar_fiscais().calendario


def test_pascoa() -> None:
    assert [pascoa(a) for a in (2008, 2016, 2024, 2025)] == [
        date(2008, 3, 23), date(2016, 3, 27), date(2024, 3, 31), date(2025, 4, 20)]  # fmt: skip


def test_carnaval_e_corpus_christi_nao_sao_dias_uteis() -> None:
    assert not dia_util(date(2024, 2, 12), CAL) and not dia_util(date(2024, 2, 13), CAL)
    assert dia_util(date(2024, 2, 14), CAL)  # quarta de cinzas
    assert not dia_util(date(2024, 5, 30), CAL)


def test_vencimento_do_darf_no_ultimo_dia_util_do_mes_seguinte() -> None:
    """RIR/2018 art. 915. Fev/2017 terminou no Carnaval (27-28/02): DARF de jan/2017 em 24/02."""
    assert vencimento_darf(2024, 1, CAL) == date(2024, 2, 29)
    assert vencimento_darf(2017, 1, CAL) == date(2017, 2, 24)
    assert vencimento_darf(2024, 4, CAL) == date(2024, 5, 31)
    assert vencimento_darf(2021, 1, CAL) == date(2021, 2, 26)  # 28/02/2021 foi domingo
    assert ultimo_dia_util(2020, 2, CAL) == date(2020, 2, 28)


@pytest.mark.parametrize(
    ("ano", "mes", "esperado"),
    [
        (2023, 1, date(2023, 1, 20)),
        (2023, 4, date(2023, 4, 20)),  # 21/04 (sexta) é Tiradentes: pregão anterior
        (2024, 11, date(2024, 11, 14)),  # 15/11 (sexta) é feriado
        (2025, 12, date(2025, 12, 19)),
        (2026, 8, date(2026, 8, 21)),
    ],
)
def test_vencimento_de_opcoes_na_terceira_sexta(ano: int, mes: int, esperado: date) -> None:
    """Regra da B3 (data/vencimentos_opcoes.csv), conferida com os exercícios reais do usuário."""
    assert vencimento_opcao(ano, mes, CAL) == esperado


@pytest.mark.parametrize(
    ("ano", "mes", "esperado"),
    [
        (2008, 4, date(2008, 4, 22)),  # 21/04: pregão seguinte
        (2010, 2, date(2010, 2, 17)),  # Carnaval: quarta de cinzas
        (2017, 11, date(2017, 11, 21)),  # 20/11, feriado de SP até 2021
        (2021, 2, date(2021, 2, 17)),
        (2021, 4, date(2021, 4, 19)),  # último vencimento na segunda
        (2021, 5, date(2021, 5, 21)),  # primeiro na sexta
        (2022, 4, date(2022, 4, 14)),
        (2026, 11, date(2026, 11, 19)),
    ],
)
def test_vencimento_antes_e_depois_de_maio_de_2021(ano: int, mes: int, esperado: date) -> None:
    """B3 (notícia de 20/04/2021): 3ª segunda (sem pregão: o seguinte) até abr/2021; 3ª sexta
    (sem pregão: o anterior) desde mai/2021. Vetores do especialista-fiscal-b3."""
    assert vencimento_opcao(ano, mes, CAL) == esperado


@pytest.mark.parametrize(
    ("ano", "mes", "semana", "esperado"),
    [
        (2024, 3, 5, date(2024, 3, 28)),  # 29/03 Sexta-feira Santa
        (2026, 5, 1, date(2026, 4, 30)),  # 01/05 feriado: recua para abril
        (2026, 12, 4, date(2026, 12, 23)),  # 25/12 e 24/12 sem pregão
        (2027, 1, 1, date(2026, 12, 30)),  # 01/01 e 31/12 (último dia útil) sem pregão
    ],
)
def test_vencimento_de_opcao_semanal(ano: int, mes: int, semana: int, esperado: date) -> None:
    """Opções semanais (B3, desde 29/01/2024): Wn vence na n-ésima sexta do mês, ou no pregão
    anterior."""
    assert vencimento_opcao(ano, mes, CAL, semana) == esperado


def test_darf_de_novembro_no_ultimo_dia_util_antes_do_fim_do_ano() -> None:
    """Res. CMN 4.880/2020 art. 3º: o último dia útil do ano não tem expediente bancário. Agenda
    RFB: 6015 de nov/2022 em 29/12/2022, de nov/2024 em 30/12/2024, de nov/2025 em 30/12/2025."""
    assert vencimento_darf(2022, 11, CAL) == date(2022, 12, 29)
    assert vencimento_darf(2024, 11, CAL) == date(2024, 12, 30)
    assert vencimento_darf(2025, 11, CAL) == date(2025, 12, 30)
    assert vencimento_darf(2017, 11, CAL) == date(2017, 12, 28)
    assert vencimento_darf(2022, 1, CAL) == date(2022, 2, 25)  # Carnaval em 28/02 e 01/03
    assert vencimento_darf(2018, 4, CAL) == date(2018, 5, 30)  # Corpus Christi em 31/05


def test_pregao_nas_excecoes_de_2014_e_2020() -> None:
    assert not dia_de_pregao(date(2014, 6, 12), CAL)  # abertura da Copa
    assert dia_de_pregao(date(2020, 7, 9), CAL) and dia_de_pregao(date(2020, 11, 20), CAL)
    assert dia_de_pregao(date(2022, 1, 25), CAL) and not dia_de_pregao(date(2021, 1, 25), CAL)


def test_ano_fora_do_calendario_falha_alto() -> None:
    with pytest.raises(ReferenciaError, match="calendário"):
        dia_util(date(2031, 1, 2), CAL)


def test_liquidacao_de_termo_no_pregao_seguinte() -> None:
    assert liquidacao_termo(date(2024, 1, 3), 60, CAL) == date(2024, 3, 4)  # 03/03 é domingo
    assert liquidacao_termo(date(2021, 1, 14), 30, CAL) == date(2021, 2, 17)  # sáb + Carnaval
    assert dia_de_pregao(date(2024, 12, 24), CAL) is False and dia_util(date(2024, 12, 24), CAL)
